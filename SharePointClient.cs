using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SharePointLongPathTool;

/// <summary>
/// Wraps SharePoint CSOM operations for scanning and remediating long file paths.
/// </summary>
public sealed class SharePointClient : IDisposable
{
    private readonly ClientContext _ctx;
    private readonly int _maxPathLength;

    private SharePointClient(ClientContext ctx, int maxPathLength)
    {
        _ctx = ctx;
        _maxPathLength = maxPathLength;
    }

    /// <summary>
    /// Connect using Azure AD app-only (client credentials) authentication.
    /// </summary>
    public static async Task<SharePointClient> ConnectAsync(
        string siteUrl, string tenantId, string clientId, string clientSecret, int maxPathLength)
    {
        var app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .Build();

        var uri = new Uri(siteUrl);
        string resource = $"{uri.Scheme}://{uri.Host}";
        var result = await app.AcquireTokenForClient(new[] { $"{resource}/.default" }).ExecuteAsync();

        var ctx = new ClientContext(siteUrl);
        ctx.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + result.AccessToken;
        };

        return new SharePointClient(ctx, maxPathLength);
    }

    /// <summary>
    /// Connect using interactive browser login (delegated permissions).
    /// </summary>
    public static async Task<SharePointClient> ConnectInteractiveAsync(
        string siteUrl, string tenantId, string clientId, int maxPathLength)
    {
        var app = PublicClientApplicationBuilder
            .Create(clientId)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .WithRedirectUri("http://localhost")
            .Build();

        var uri = new Uri(siteUrl);
        string resource = $"{uri.Scheme}://{uri.Host}";
        var result = await app.AcquireTokenInteractive(new[] { $"{resource}/AllSites.Manage" })
            .ExecuteAsync();

        var ctx = new ClientContext(siteUrl);
        ctx.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + result.AccessToken;
        };

        return new SharePointClient(ctx, maxPathLength);
    }

    /// <summary>
    /// Recursively scans all files in the given library and returns those whose
    /// full server-relative URL exceeds the configured max path length.
    /// </summary>
    public async Task<List<LongPathItem>> ScanAsync(string libraryTitle)
    {
        var web = _ctx.Web;
        var list = web.Lists.GetByTitle(libraryTitle);
        _ctx.Load(web, w => w.ServerRelativeUrl);
        _ctx.Load(list, l => l.RootFolder.ServerRelativeUrl);
        await _ctx.ExecuteQueryAsync();

        string siteUrl = _ctx.Url.TrimEnd('/');
        var results = new List<LongPathItem>();
        var caml = new CamlQuery
        {
            ViewXml = @"<View Scope='RecursiveAll'>
                            <Query/>
                            <RowLimit>5000</RowLimit>
                         </View>"
        };

        ListItemCollection items;
        do
        {
            items = list.GetItems(caml);
            _ctx.Load(items, ic => ic.ListItemCollectionPosition,
                ic => ic.Include(
                    i => i.FileSystemObjectType,
                    i => i["FileRef"],
                    i => i["FileLeafRef"],
                    i => i.File.Length));
            await _ctx.ExecuteQueryAsync();

            foreach (var item in items)
            {
                string serverRelativeUrl = item["FileRef"]?.ToString() ?? "";
                string fullPath = siteUrl + serverRelativeUrl;

                if (fullPath.Length > _maxPathLength)
                {
                    results.Add(new LongPathItem
                    {
                        FullPath = fullPath,
                        ServerRelativeUrl = serverRelativeUrl,
                        FileName = item["FileLeafRef"]?.ToString() ?? "",
                        PathLength = fullPath.Length,
                        IsFolder = item.FileSystemObjectType == FileSystemObjectType.Folder,
                        FileSize = item.FileSystemObjectType == FileSystemObjectType.File
                            ? Convert.ToInt64(item.File.Length)
                            : 0
                    });
                }
            }

            caml.ListItemCollectionPosition = items.ListItemCollectionPosition;
        } while (items.ListItemCollectionPosition != null);

        return results.OrderByDescending(r => r.PathLength).ToList();
    }

    /// <summary>
    /// Shortens long paths by truncating folder or file names that push the full
    /// path beyond the limit. Folders are renamed first (deepest first) to maximise
    /// the effect for nested items.
    /// </summary>
    public async Task<List<RemediationResult>> RemediateAsync(
        string libraryTitle, bool dryRun, int? truncateLength)
    {
        var longPaths = await ScanAsync(libraryTitle);
        if (longPaths.Count == 0)
            return new List<RemediationResult>();

        int targetNameLen = truncateLength ?? 50;
        var results = new List<RemediationResult>();
        var renamedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Process folders first (deepest first), then files
        var folders = longPaths.Where(p => p.IsFolder)
            .OrderByDescending(p => p.ServerRelativeUrl.Count(c => c == '/'))
            .ToList();
        var files = longPaths.Where(p => !p.IsFolder).ToList();

        foreach (var item in folders.Concat(files))
        {
            if (renamedPaths.Contains(item.ServerRelativeUrl))
                continue;

            string originalName = item.FileName;
            string newName = ShortenName(originalName, targetNameLen, item.IsFolder);

            if (newName == originalName)
                continue;

            var result = new RemediationResult
            {
                OriginalPath = item.ServerRelativeUrl,
                OriginalName = originalName,
                NewName = newName,
                DryRun = dryRun
            };

            if (!dryRun)
            {
                try
                {
                    if (item.IsFolder)
                    {
                        var folder = _ctx.Web.GetFolderByServerRelativeUrl(item.ServerRelativeUrl);
                        folder.MoveTo(item.ServerRelativeUrl.Replace(originalName, newName));
                        await _ctx.ExecuteQueryAsync();
                    }
                    else
                    {
                        var file = _ctx.Web.GetFileByServerRelativeUrl(item.ServerRelativeUrl);
                        string destUrl = item.ServerRelativeUrl
                            [..item.ServerRelativeUrl.LastIndexOf('/')]
                            + "/" + newName;
                        file.MoveTo(destUrl, MoveOperations.Overwrite);
                        await _ctx.ExecuteQueryAsync();
                    }

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Error = ex.Message;
                }
            }
            else
            {
                result.Success = true;
            }

            results.Add(result);
            renamedPaths.Add(item.ServerRelativeUrl);
        }

        return results;
    }

    /// <summary>
    /// Truncates a file/folder name while preserving the extension for files.
    /// </summary>
    internal static string ShortenName(string name, int maxNameLength, bool isFolder)
    {
        if (name.Length <= maxNameLength)
            return name;

        if (isFolder)
            return name[..maxNameLength].TrimEnd(' ', '.');

        string ext = Path.GetExtension(name);
        string stem = Path.GetFileNameWithoutExtension(name);
        int allowedStem = maxNameLength - ext.Length;

        if (allowedStem < 5)
            allowedStem = 5;

        return stem[..Math.Min(stem.Length, allowedStem)].TrimEnd(' ', '.') + ext;
    }

    public void Dispose() => _ctx.Dispose();
}

public class LongPathItem
{
    public string FullPath { get; set; } = "";
    public string ServerRelativeUrl { get; set; } = "";
    public string FileName { get; set; } = "";
    public int PathLength { get; set; }
    public bool IsFolder { get; set; }
    public long FileSize { get; set; }
}

public class RemediationResult
{
    public string OriginalPath { get; set; } = "";
    public string OriginalName { get; set; } = "";
    public string NewName { get; set; } = "";
    public bool DryRun { get; set; }
    public bool Success { get; set; }
    public string? Error { get; set; }
}
