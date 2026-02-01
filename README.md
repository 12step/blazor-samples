# SharePoint Long Path Tool

A .NET 8 console app that scans a SharePoint document library for file/folder paths exceeding 400 characters (the OneDrive sync limit) and optionally remediates them by shortening names.

## Prerequisites

- .NET 8 SDK
- An Azure AD app registration with **Sites.Read.All** (scan) or **Sites.Manage.All** (fix) permissions
- SharePoint Online site URL

## Setup

```bash
cd SharePointLongPathTool
dotnet restore
dotnet build
```

## Usage

### Scan – export long paths to CSV

```bash
dotnet run -- scan \
  --site-url "https://contoso.sharepoint.com/sites/team" \
  --tenant-id "your-tenant-id" \
  --client-id "your-client-id" \
  --client-secret "your-secret" \
  --library "Documents" \
  --max-path 400 \
  --output long-paths.csv
```

Omit `--client-secret` to use interactive browser login instead.

### Fix – shorten names (dry run first)

```bash
# Preview changes
dotnet run -- fix \
  --site-url "https://contoso.sharepoint.com/sites/team" \
  --tenant-id "your-tenant-id" \
  --client-id "your-client-id" \
  --client-secret "your-secret" \
  --dry-run true

# Apply changes
dotnet run -- fix \
  --site-url "https://contoso.sharepoint.com/sites/team" \
  --tenant-id "your-tenant-id" \
  --client-id "your-client-id" \
  --client-secret "your-secret" \
  --dry-run false \
  --name-max 50
```

### Options

| Option | Default | Description |
|---|---|---|
| `--site-url` | *(required)* | SharePoint site URL |
| `--library` | `Documents` | Document library title |
| `--tenant-id` | *(required)* | Azure AD tenant ID |
| `--client-id` | *(required)* | Azure AD app client ID |
| `--client-secret` | *(optional)* | Client secret; omit for interactive login |
| `--max-path` | `400` | Path length threshold |
| `--output` | `long-paths.csv` | CSV output path (scan only) |
| `--dry-run` | `true` | Preview without renaming (fix only) |
| `--name-max` | `50` | Max characters per file/folder name (fix only) |
