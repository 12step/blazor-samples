using System.CommandLine;
using System.Text;
using SharePointLongPathTool;

var siteUrlOption = new Option<string>("--site-url", "SharePoint site URL") { IsRequired = true };
var libraryOption = new Option<string>("--library", () => "Documents", "Document library title");
var tenantIdOption = new Option<string>("--tenant-id", "Azure AD tenant ID") { IsRequired = true };
var clientIdOption = new Option<string>("--client-id", "Azure AD app (client) ID") { IsRequired = true };
var clientSecretOption = new Option<string?>("--client-secret", "Client secret (omit for interactive login)");
var maxPathOption = new Option<int>("--max-path", () => 400, "Maximum allowed path length");
var outputOption = new Option<string>("--output", () => "long-paths.csv", "Output CSV file path");

// ── scan command ──────────────────────────────────────────────────────────
var scanCommand = new Command("scan", "Scan a SharePoint library and export long paths to CSV");
scanCommand.AddOption(siteUrlOption);
scanCommand.AddOption(libraryOption);
scanCommand.AddOption(tenantIdOption);
scanCommand.AddOption(clientIdOption);
scanCommand.AddOption(clientSecretOption);
scanCommand.AddOption(maxPathOption);
scanCommand.AddOption(outputOption);

scanCommand.SetHandler(async (string siteUrl, string library, string tenantId,
    string clientId, string? clientSecret, int maxPath, string output) =>
{
    using var client = await Connect(siteUrl, tenantId, clientId, clientSecret, maxPath);

    Console.WriteLine($"Scanning \"{library}\" for paths longer than {maxPath} characters...");
    var items = await client.ScanAsync(library);

    if (items.Count == 0)
    {
        Console.WriteLine("No long paths found.");
        return;
    }

    var sb = new StringBuilder();
    sb.AppendLine("FullPath,ServerRelativeUrl,FileName,PathLength,IsFolder,FileSize");
    foreach (var item in items)
    {
        sb.AppendLine($"\"{Esc(item.FullPath)}\",\"{Esc(item.ServerRelativeUrl)}\",\"{Esc(item.FileName)}\",{item.PathLength},{item.IsFolder},{item.FileSize}");
    }

    await File.WriteAllTextAsync(output, sb.ToString());
    Console.WriteLine($"Found {items.Count} items exceeding {maxPath} characters.");
    Console.WriteLine($"Report saved to: {Path.GetFullPath(output)}");

}, siteUrlOption, libraryOption, tenantIdOption, clientIdOption,
   clientSecretOption, maxPathOption, outputOption);

// ── fix command ───────────────────────────────────────────────────────────
var dryRunOption = new Option<bool>("--dry-run", () => true, "Preview changes without renaming");
var truncateLenOption = new Option<int>("--name-max", () => 50, "Max characters for a single file/folder name");

var fixCommand = new Command("fix", "Shorten file and folder names that cause long paths");
fixCommand.AddOption(siteUrlOption);
fixCommand.AddOption(libraryOption);
fixCommand.AddOption(tenantIdOption);
fixCommand.AddOption(clientIdOption);
fixCommand.AddOption(clientSecretOption);
fixCommand.AddOption(maxPathOption);
fixCommand.AddOption(dryRunOption);
fixCommand.AddOption(truncateLenOption);

fixCommand.SetHandler(async (string siteUrl, string library, string tenantId,
    string clientId, string? clientSecret, int maxPath, bool dryRun, int nameMax) =>
{
    using var client = await Connect(siteUrl, tenantId, clientId, clientSecret, maxPath);

    string mode = dryRun ? "DRY RUN" : "LIVE";
    Console.WriteLine($"[{mode}] Remediating long paths in \"{library}\"...");

    var results = await client.RemediateAsync(library, dryRun, nameMax);

    if (results.Count == 0)
    {
        Console.WriteLine("Nothing to remediate.");
        return;
    }

    foreach (var r in results)
    {
        string status = r.Success ? "OK" : $"FAIL: {r.Error}";
        Console.WriteLine($"  [{status}] {r.OriginalName} -> {r.NewName}");
        Console.WriteLine($"           {r.OriginalPath}");
    }

    Console.WriteLine($"\n{results.Count} items processed. Succeeded: {results.Count(r => r.Success)}, Failed: {results.Count(r => !r.Success)}");

    if (dryRun)
        Console.WriteLine("\nThis was a dry run. Re-run with --dry-run false to apply changes.");

}, siteUrlOption, libraryOption, tenantIdOption, clientIdOption,
   clientSecretOption, maxPathOption, dryRunOption, truncateLenOption);

// ── root ──────────────────────────────────────────────────────────────────
var rootCommand = new RootCommand("SharePoint Long Path Tool – find and fix paths that exceed OneDrive sync limits");
rootCommand.AddCommand(scanCommand);
rootCommand.AddCommand(fixCommand);

return await rootCommand.InvokeAsync(args);

// ── helpers ───────────────────────────────────────────────────────────────
static async Task<SharePointClient> Connect(
    string siteUrl, string tenantId, string clientId, string? clientSecret, int maxPath)
{
    if (!string.IsNullOrEmpty(clientSecret))
        return await SharePointClient.ConnectAsync(siteUrl, tenantId, clientId, clientSecret, maxPath);

    Console.WriteLine("No client secret provided – launching interactive browser login...");
    return await SharePointClient.ConnectInteractiveAsync(siteUrl, tenantId, clientId, maxPath);
}

static string Esc(string value) => value.Replace("\"", "\"\"");
