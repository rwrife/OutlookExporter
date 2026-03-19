using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

internal sealed class Program
{
    private static int Main(string[] args)
    {
        var options = ParseArgs(args);

        try
        {
            Directory.CreateDirectory(options.OutputFolder);

            Outlook.Application outlookApp = GetOutlookApplication();
            Outlook.NameSpace session = outlookApp.Session;
            session.Logon(Missing.Value, Missing.Value, false, false);

            Outlook.MAPIFolder rootFolder = ResolveFolder(session, options.FolderPath);
            Console.WriteLine($"Exporting from: {rootFolder.FolderPath}");

            int exported = 0;
            TraverseFolder(rootFolder, options, ref exported);

            Console.WriteLine();
            Console.WriteLine($"Done. Exported {exported} messages to:");
            Console.WriteLine(Path.GetFullPath(options.OutputFolder));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Export failed:");
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static Outlook.Application GetOutlookApplication()
    {
        try
        {
            var clsid = Type.GetTypeFromProgID("Outlook.Application")!.GUID;
            Marshal.ThrowExceptionForHR(GetActiveObject(clsid, IntPtr.Zero, out object obj));
            return (Outlook.Application)obj;
        }
        catch
        {
            var app = new Outlook.Application();
            var explorer = app.Explorers.Count > 0
                ? app.Explorers[1]
                : app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).GetExplorer();
            explorer.Display();
            return app;
        }
    }

    private static void TraverseFolder(Outlook.MAPIFolder folder, ExportOptions options, ref int exported)
    {
        ExportFolderMessages(folder, options, ref exported);

        if (!options.RecurseSubfolders)
        {
            return;
        }

        foreach (Outlook.MAPIFolder subFolder in folder.Folders)
        {
            TraverseFolder(subFolder, options, ref exported);
        }
    }

    private static void ExportFolderMessages(Outlook.MAPIFolder folder, ExportOptions options, ref int exported)
    {
        Outlook.Items items = folder.Items;
        items.Sort("[ReceivedTime]", true);

        int processed = 0;

        foreach (object item in items)
        {
            if (options.MaxMessages.HasValue && exported >= options.MaxMessages.Value)
            {
                return;
            }

            if (item is not Outlook.MailItem mail)
            {
                continue;
            }

            processed++;
            if (options.SinceUtc.HasValue && mail.ReceivedTime.ToUniversalTime() < options.SinceUtc.Value)
            {
                continue;
            }

            var dto = MailExport.FromMailItem(mail, folder.FolderPath, options.IncludeHtmlBody);
            var safeName = BuildSafeFileName(dto);
            var outputPath = Path.Combine(options.OutputFolder, safeName);

            var json = JsonSerializer.Serialize(dto, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(outputPath, json);
            exported++;

            Console.WriteLine($"[{exported}] {dto.ReceivedTimeUtc:u} | {dto.Subject}");
        }
    }

    private static string BuildSafeFileName(MailExport dto)
    {
        var subject = string.IsNullOrWhiteSpace(dto.Subject) ? "no-subject" : dto.Subject.Trim();
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            subject = subject.Replace(c, '_');
        }

        if (subject.Length > 80)
        {
            subject = subject[..80];
        }

        var timestamp = dto.ReceivedTimeUtc?.ToString("yyyyMMdd_HHmmss") ?? "unknown-date";
        var idPart = string.IsNullOrWhiteSpace(dto.EntryId) ? Guid.NewGuid().ToString("N")[..8] : dto.EntryId.GetHashCode().ToString("X8");
        return $"{timestamp}_{subject}_{idPart}.json";
    }

    private static Outlook.MAPIFolder ResolveFolder(Outlook.NameSpace session, string folderPath)
    {
        // Examples:
        //   "\\user@company.com\\Inbox"
        //   "\\Mailbox - Ryan Rife\\Inbox"
        //   "Inbox" (uses default Inbox)
        if (string.Equals(folderPath, "Inbox", StringComparison.OrdinalIgnoreCase))
        {
            return session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        }

        if (!folderPath.StartsWith("\\"))
        {
            throw new ArgumentException("Folder path must start with '\\\\' or be exactly 'Inbox'.");
        }

        var parts = folderPath.Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            throw new ArgumentException("Folder path is empty.");
        }

        Outlook.MAPIFolder current = session.Folders[parts[0]];
        for (int i = 1; i < parts.Length; i++)
        {
            current = current.Folders[parts[i]];
        }

        return current;
    }

    private static ExportOptions ParseArgs(string[] args)
    {
        var options = new ExportOptions();

        foreach (string arg in args)
        {
            if (arg.StartsWith("--out=", StringComparison.OrdinalIgnoreCase))
            {
                options.OutputFolder = arg["--out=".Length..].Trim('"');
            }
            else if (arg.StartsWith("--folder=", StringComparison.OrdinalIgnoreCase))
            {
                options.FolderPath = arg["--folder=".Length..].Trim('"');
            }
            else if (arg.StartsWith("--max=", StringComparison.OrdinalIgnoreCase))
            {
                var maxVal = arg["--max=".Length..];
                if (int.TryParse(maxVal, out int maxMessages))
                {
                    options.MaxMessages = maxMessages;
                }
            }
            else if (arg.StartsWith("--sinceUtc=", StringComparison.OrdinalIgnoreCase))
            {
                options.SinceUtc = DateTime.Parse(arg["--sinceUtc=".Length..], null, System.Globalization.DateTimeStyles.AdjustToUniversal);
            }
            else if (arg.Equals("--no-recurse", StringComparison.OrdinalIgnoreCase))
            {
                options.RecurseSubfolders = false;
            }
            else if (arg.Equals("--include-html", StringComparison.OrdinalIgnoreCase))
            {
                options.IncludeHtmlBody = true;
            }
            else if (arg is "--help" or "-h" or "/?")
            {
                PrintHelpAndExit();
            }
        }

        return options;
    }

    private static void PrintHelpAndExit()
    {
        Console.WriteLine("""
Outlook COM Reader POC

Examples:
  OutlookComReader.exe --folder=Inbox --max=100
  OutlookComReader.exe --folder="\\Mailbox - Ryan Rife\\Inbox" --max=250 --out=export
  OutlookComReader.exe --folder=Inbox --sinceUtc=2026-03-01T00:00:00Z --include-html

Options:
  --folder=PATH        Outlook folder path. Use Inbox for default Inbox.
  --out=PATH           Output folder. Default: exported-json
  --max=N              Maximum number of messages to export.
  --sinceUtc=ISO8601   Only export messages received on/after this UTC timestamp.
  --no-recurse         Do not recurse into subfolders.
  --include-html       Include HTMLBody in the JSON output.
""");
        Environment.Exit(0);
    }
}

internal sealed class ExportOptions
{
    public string FolderPath { get; set; } = "Inbox";
    public string OutputFolder { get; set; } = "exported-json";
    public int? MaxMessages { get; set; }
    public DateTime? SinceUtc { get; set; }
    public bool RecurseSubfolders { get; set; } = true;
    public bool IncludeHtmlBody { get; set; }
}

internal sealed class MailExport
{
    public string? EntryId { get; init; }
    public string? StoreId { get; init; }
    public string? FolderPath { get; init; }
    public string? Subject { get; init; }
    public string? SenderName { get; init; }
    public string? SenderEmailAddress { get; init; }
    public string? To { get; init; }
    public string? Cc { get; init; }
    public string? Bcc { get; init; }
    public DateTime? ReceivedTimeUtc { get; init; }
    public DateTime? SentOnUtc { get; init; }
    public string? Body { get; init; }
    public string? HtmlBody { get; init; }
    public string? ConversationTopic { get; init; }
    public string? ConversationId { get; init; }
    public string? MessageClass { get; init; }
    public bool UnRead { get; init; }
    public bool HasAttachments { get; init; }
    public string? OutlookClassicLink { get; init; }
    public string? OutlookWebLink { get; init; }

    public static MailExport FromMailItem(Outlook.MailItem mail, string folderPath, bool includeHtmlBody)
    {
        string? conversationId = null;
        try
        {
            conversationId = mail.ConversationID;
        }
        catch
        {
            // Some stores/items may not expose this consistently.
        }

        return new MailExport
        {
            EntryId = mail.EntryID,
            StoreId = mail.Parent is Outlook.MAPIFolder f ? f.StoreID : null,
            FolderPath = folderPath,
            Subject = mail.Subject,
            SenderName = mail.SenderName,
            SenderEmailAddress = mail.SenderEmailAddress,
            To = mail.To,
            Cc = mail.CC,
            Bcc = mail.BCC,
            ReceivedTimeUtc = SafeToUtc(mail.ReceivedTime),
            SentOnUtc = SafeToUtc(mail.SentOn),
            Body = mail.Body,
            HtmlBody = includeHtmlBody ? mail.HTMLBody : null,
            ConversationTopic = mail.ConversationTopic,
            ConversationId = conversationId,
            MessageClass = mail.MessageClass,
            UnRead = mail.UnRead,
            HasAttachments = mail.Attachments.Count > 0,
            OutlookClassicLink = BuildOutlookClassicLink(mail.EntryID),
            OutlookWebLink = BuildOutlookWebLink(mail.EntryID)
        };
    }

    private static string? BuildOutlookClassicLink(string? entryId)
    {
        if (string.IsNullOrEmpty(entryId)) return null;
        return $"outlook:{entryId}";
    }

    private static string? BuildOutlookWebLink(string? entryId)
    {
        if (string.IsNullOrEmpty(entryId)) return null;
        try
        {
            var bytes = Convert.FromHexString(entryId);
            var base64 = Convert.ToBase64String(bytes);
            return $"https://outlook.office.com/owa/?ItemID={Uri.EscapeDataString(base64)}&exvsurl=1&viewmodel=ReadMessageItem";
        }
        catch
        {
            return null;
        }
    }

    private static DateTime? SafeToUtc(DateTime dt)
    {
        if (dt == DateTime.MinValue)
        {
            return null;
        }

        return dt.Kind == DateTimeKind.Utc ? dt : dt.ToUniversalTime();
    }
}
