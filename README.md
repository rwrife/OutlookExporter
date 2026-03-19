# Outlook COM Reader POC

`OutlookComReader` is a Windows-only .NET 8 console application that connects to classic Outlook through COM interop and exports email messages as JSON files.

This repository is intended as a proof of concept for mailbox export scenarios where Outlook is already installed and signed in on the machine running the tool.

## What it does

- Connects to a running classic Outlook instance, or starts Outlook if needed
- Resolves either the default `Inbox` or a fully qualified Outlook folder path
- Walks the selected folder and, by default, its subfolders
- Exports one JSON file per mail item
- Optionally filters by received date and includes the HTML body

## JSON fields

Each exported message can include:

- Outlook identifiers: `EntryId`, `StoreId`
- Mailbox metadata: `FolderPath`, `ConversationTopic`, `ConversationId`, `MessageClass`
- Sender and recipient data: `SenderName`, `SenderEmailAddress`, `To`, `Cc`, `Bcc`
- Timestamps: `ReceivedTimeUtc`, `SentOnUtc`
- Content: `Body`, optional `HtmlBody`
- Flags and links: `UnRead`, `HasAttachments`, `OutlookClassicLink`, `OutlookWebLink`

## Requirements

- Windows
- Classic Outlook for Windows installed
- Outlook profile configured and able to open the mailbox you want to export
- .NET 8 SDK
- A local desktop session that can launch Outlook if it is not already running

## Important constraints

- This project targets `net8.0-windows` and uses the Outlook COM object library.
- It is not cross-platform.
- It is not suitable for standard GitHub Actions hosted runners because they do not provide an interactive, configured Outlook desktop environment.
- If you publish this on GitHub, treat it as a locally run desktop utility, not a cloud automation job.

## Repository contents

- `Program.cs`: app entry point, argument parsing, Outlook access, and JSON export logic
- `OutlookComReader.csproj`: .NET 8 project file with the Outlook COM reference
- `outlook-com-reader-poc.sln`: Visual Studio solution

## Build locally

### Visual Studio

1. Open `outlook-com-reader-poc.sln` in Visual Studio on Windows.
2. Restore and build the solution.
3. Make sure Outlook opens successfully under the same Windows user.

### .NET CLI

Run from the repository root:

```powershell
dotnet build .\OutlookComReader.csproj -c Debug
```

Note: because this project uses Outlook COM interop, command-line build support depends on the machine having Outlook installed correctly.

## Run locally

From the repository root:

```powershell
dotnet run --project .\OutlookComReader.csproj -- --folder=Inbox --max=100
```

Or run the compiled executable from `bin\Debug\net8.0-windows\`.

## Command-line options

- `--folder=PATH`: Outlook folder path. Use `Inbox` for the default Inbox.
- `--out=PATH`: Output directory. Defaults to `exported-json`.
- `--max=N`: Maximum number of messages to export.
- `--sinceUtc=ISO8601`: Export only messages received on or after a UTC timestamp.
- `--no-recurse`: Do not export subfolders.
- `--include-html`: Include `HTMLBody` in the JSON output.
- `--help`: Print usage help.

## Examples

```powershell
# Default Inbox, first 100 messages
dotnet run --project .\OutlookComReader.csproj -- --folder=Inbox --max=100

# Specific mailbox folder
dotnet run --project .\OutlookComReader.csproj -- --folder="\\Mailbox - Ryan Rife\\Inbox" --max=250

# Export newer messages only, include HTML
dotnet run --project .\OutlookComReader.csproj -- --folder=Inbox --sinceUtc=2026-03-01T00:00:00Z --include-html

# Write output to a custom folder without recursing into subfolders
dotnet run --project .\OutlookComReader.csproj -- --folder=Inbox --out=.\exports --no-recurse
```

## Output

By default, exported files are written to:

```text
exported-json
```

File names are generated from the received timestamp, a sanitized subject, and a message-derived identifier.

## Using this project on GitHub

If your goal is to store and share this project in a GitHub repository:

1. Commit the source files, solution file, and README.
2. Do not commit `bin/`, `obj/`, or exported mailbox JSON data.
3. Use the included `.gitignore` so local build artifacts and message exports stay out of the repo.
4. Document clearly that contributors must run the tool on Windows with classic Outlook installed.

If your goal is to run this from GitHub Actions:

- That is not a good fit for the current implementation.
- GitHub-hosted runners do not provide a signed-in Outlook desktop profile.
- A self-hosted Windows runner with Outlook installed and configured would be required, but even then COM desktop automation in CI is fragile.

## Suggested GitHub workflow

1. Clone the repository on a Windows workstation.
2. Open Outlook and verify the mailbox is available.
3. Build the project locally.
4. Run the exporter locally.
5. Review the generated JSON outside the repository, or in an ignored output folder.

## Notes

- This is for classic Outlook only.
- The app attempts to attach to an already running Outlook instance first.
- Some mailbox stores may behave differently for fields such as `ConversationID` or deep folder access.
- Exported message data can contain sensitive information, so keep output directories out of source control.
