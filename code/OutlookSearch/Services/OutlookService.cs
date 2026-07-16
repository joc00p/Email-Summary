using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace OutlookSearch.Services;

public record MailFolderNode(string EntryId, string StoreId, string Name, List<MailFolderNode> Children);

public record SearchOptions(
    string? Keyword,
    string? Subject,
    string? From,
    DateTime? DateFrom,
    DateTime? DateTo);

/// <summary>Result of a search, including whether it was cut short so the UI can be honest about it.</summary>
public record SearchOutcome(List<EmailResult> Items, bool Truncated, int FoldersSearched, int FoldersFailed);

public class EmailResult
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public string Subject { get; init; } = "";
    public string From { get; init; } = "";
    public string To { get; init; } = "";
    public DateTime? ReceivedAt { get; init; }
    public string FolderName { get; init; } = "";
    public string BodyPreview { get; init; } = "";
    public bool HasAttachments { get; init; }
}

/// <summary>
/// Talks to the locally installed Outlook via late-bound COM (the Outlook Object
/// Model). No NuGet packages, no Azure — it attaches to the user's running Outlook
/// profile and can read any mailbox already open in it (own + shared/delegated).
/// </summary>
public class OutlookService : IDisposable
{
    // Outlook enum constants (avoids needing the interop assembly).
    private const int olFolderInbox = 6;
    private const int MaxResultsDefault = 2000;

    // MAPI property tags for the sender's real SMTP address. SenderEmailAddress is an
    // X.500 DN (not SMTP) for internal Exchange senders, so we read these instead.
    private const string PrSenderSmtp  = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
    private const string PrSentRepSmtp = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F";

    private readonly StaTaskScheduler _sta = new();
    private dynamic? _app;
    private dynamic? _ns;

    /// <summary>The result cap applied per search; when hit, the outcome is flagged as truncated.</summary>
    public int MaxResults { get; } = MaxResultsDefault;

    private void EnsureApp()
    {
        if (_app != null) return;
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Outlook is not installed on this machine (Outlook.Application COM class not found).");
        _app = Activator.CreateInstance(type)!;   // attaches to running Outlook, or starts it
        _ns = _app!.GetNamespace("MAPI");
    }

    /// <summary>Every mailbox / store in the current Outlook profile, as a checkable folder tree.</summary>
    public Task<List<MailFolderNode>> GetMailboxTreesAsync() => _sta.Run(() =>
    {
        EnsureApp();
        var roots = new List<MailFolderNode>();
        dynamic? stores = null;
        try
        {
            stores = _ns!.Stores;
            int count = stores.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? store = null;
                dynamic? root = null;
                try
                {
                    store = stores.Item(i);
                    string storeId = store.StoreID;
                    root = store.GetRootFolder();
                    roots.Add(BuildNode(root, storeId));
                }
                catch { /* skip stores we cannot open */ }
                finally { Rel(root); Rel(store); }
            }
        }
        finally { Rel(stores); }
        return roots;
    });

    /// <summary>
    /// Resolve a person by name or email against Outlook's address book (the org GAL / AD)
    /// and return their mailbox folder tree — works when you have delegate/shared access.
    /// </summary>
    public Task<MailFolderNode?> OpenSharedMailboxAsync(string nameOrEmail) => _sta.Run<MailFolderNode?>(() =>
    {
        EnsureApp();
        dynamic? recipient = null;
        dynamic? inbox = null;
        dynamic? root = null;
        try
        {
            recipient = _ns!.CreateRecipient(nameOrEmail);
            recipient.Resolve();
            if (!(bool)recipient.Resolved)
                throw new InvalidOperationException($"Could not resolve '{nameOrEmail}' in the address book.");

            // GetSharedDefaultFolder gives us their Inbox; its parent is the mailbox root,
            // from which we can enumerate the whole (accessible) folder tree.
            inbox = _ns.GetSharedDefaultFolder(recipient, olFolderInbox);
            try { root = inbox.Parent; }
            catch { root = inbox; }

            string storeId = "";
            try { storeId = root.StoreID; } catch { }
            return BuildNode(root, storeId);
        }
        finally
        {
            if (!ReferenceEquals(root, inbox)) Rel(root);
            Rel(inbox);
            Rel(recipient);
        }
    });

    private static MailFolderNode BuildNode(dynamic folder, string storeId)
    {
        var children = new List<MailFolderNode>();
        dynamic? subs = null;
        try
        {
            subs = folder.Folders;
            int c = subs.Count;
            for (int i = 1; i <= c; i++)
            {
                dynamic? child = null;
                try { child = subs.Item(i); children.Add(BuildNode(child, storeId)); }
                catch { /* skip folders we lack permission to enumerate */ }
                finally { Rel(child); }
            }
        }
        catch { }
        finally { Rel(subs); }

        string name = "Folder"; try { name = folder.Name; } catch { }
        string eid = "";  try { eid = folder.EntryID; } catch { }
        string sid = storeId; try { if (string.IsNullOrEmpty(sid)) sid = folder.StoreID; } catch { }

        return new MailFolderNode(eid, sid, name, children);
    }

    public Task<SearchOutcome> SearchAsync(
        IReadOnlyList<(string EntryId, string StoreId, string Name)> folders,
        SearchOptions opts,
        CancellationToken ct) => _sta.Run(() =>
    {
        EnsureApp();
        var dasl = BuildDasl(opts);
        var results = new List<EmailResult>();
        bool truncated = false;
        int foldersSearched = 0, foldersFailed = 0;

        foreach (var (entryId, storeId, folderName) in folders)
        {
            if (ct.IsCancellationRequested) break;
            if (string.IsNullOrEmpty(entryId)) continue;

            dynamic? folder = null;
            dynamic? items = null;
            dynamic? candidates = null;
            try
            {
                try { folder = string.IsNullOrEmpty(storeId) ? _ns!.GetFolderFromID(entryId) : _ns!.GetFolderFromID(entryId, storeId); }
                catch { foldersFailed++; continue; }

                try { items = folder.Items; } catch { foldersFailed++; continue; }

                // Restrict narrows the set store-side for speed; fall back to the full
                // collection if the DASL query is rejected. Match() below is authoritative.
                candidates = items;
                bool restrictApplied = false;
                if (!string.IsNullOrEmpty(dasl))
                {
                    try { candidates = items.Restrict(dasl); restrictApplied = true; }
                    catch { candidates = items; restrictApplied = false; }
                }
                try { candidates.Sort("[ReceivedTime]", true); } catch { /* order enforced again below */ }

                int n;
                try { n = candidates.Count; } catch { foldersFailed++; continue; }
                foldersSearched++;

                for (int i = 1; i <= n; i++)
                {
                    if (ct.IsCancellationRequested) break;
                    if (results.Count >= MaxResults) { truncated = true; break; }

                    dynamic? item = null;
                    try
                    {
                        try { item = candidates.Item(i); } catch { continue; }

                        string cls = "";
                        try { cls = item.MessageClass; } catch { }
                        if (!cls.StartsWith("IPM.Note")) continue;   // mail items only

                        if (!Match(item, opts, restrictApplied)) continue;
                        results.Add(MapItem(item, storeId, folderName));
                    }
                    catch { /* skip any item that won't map cleanly */ }
                    finally { Rel(item); }
                }
            }
            finally
            {
                if (!ReferenceEquals(candidates, items)) Rel(candidates);
                Rel(items);
                Rel(folder);
            }

            if (truncated) break;   // global cap reached; remaining folders not searched
        }

        // Final ordering is authoritative even if a folder's COM Sort failed.
        var ordered = results
            .OrderByDescending(r => r.ReceivedAt ?? DateTime.MinValue)
            .ToList();
        return new SearchOutcome(ordered, truncated, foldersSearched, foldersFailed);
    });

    public Task<string> GetBodyAsync(string entryId, string storeId) => _sta.Run(() =>
    {
        EnsureApp();
        dynamic? item = null;
        try
        {
            item = string.IsNullOrEmpty(storeId)
                ? _ns!.GetItemFromID(entryId)
                : _ns!.GetItemFromID(entryId, storeId);

            string body = SafeStr(() => item!.Body);
            if (!string.IsNullOrWhiteSpace(body)) return body;

            // Plain body empty (e.g. HTML-only mail) — fall back to the HTML, stripped.
            string html = SafeStr(() => item!.HTMLBody);
            return StripHtml(html);
        }
        catch { return ""; }
        finally { Rel(item); }
    });

    // In-memory predicate — the source of truth for whether an item matches.
    // When <paramref name="restrictApplied"/> is true the store-side DASL filter already
    // matched keyword against subject+body, so we skip the expensive full-body re-read.
    private static bool Match(dynamic item, SearchOptions o, bool restrictApplied)
    {
        try
        {
            DateTime? received = TryGetDate(item);
            if (o.DateFrom.HasValue && (received is null || received.Value < o.DateFrom.Value.Date))
                return false;
            if (o.DateTo.HasValue && (received is null || received.Value >= o.DateTo.Value.Date.AddDays(1)))
                return false;

            if (!string.IsNullOrEmpty(o.Subject))
            {
                string subj = SafeStr(() => item.Subject);
                if (subj.IndexOf(o.Subject, StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (!string.IsNullOrEmpty(o.From))
            {
                string sn = SafeStr(() => item.SenderName);
                string smtp = GetSenderSmtp(item);
                if (sn.IndexOf(o.From, StringComparison.OrdinalIgnoreCase) < 0 &&
                    smtp.IndexOf(o.From, StringComparison.OrdinalIgnoreCase) < 0)
                    return false;
            }

            if (!string.IsNullOrEmpty(o.Keyword) && !restrictApplied)
            {
                // Only reached on the fallback path (Restrict was rejected) — reading the
                // body here is unavoidable to honor the keyword.
                string subj = SafeStr(() => item.Subject);
                string body = SafeStr(() => item.Body);
                if (subj.IndexOf(o.Keyword, StringComparison.OrdinalIgnoreCase) < 0 &&
                    body.IndexOf(o.Keyword, StringComparison.OrdinalIgnoreCase) < 0)
                    return false;
            }

            return true;
        }
        catch { return false; }
    }

    // Returns the sender's real SMTP address (works for both Exchange and internet senders).
    private static string GetSenderSmtp(dynamic item)
    {
        foreach (var tag in new[] { PrSenderSmtp, PrSentRepSmtp })
        {
            try
            {
                string v = item.PropertyAccessor.GetProperty(tag) as string ?? "";
                if (v.Contains('@')) return v;
            }
            catch { }
        }
        // Fallback: SenderEmailAddress is a real SMTP address only for non-Exchange senders.
        try
        {
            if (!string.Equals(SafeStr(() => item.SenderEmailType), "EX", StringComparison.OrdinalIgnoreCase))
            {
                string addr = SafeStr(() => item.SenderEmailAddress);
                if (addr.Contains('@')) return addr;
            }
        }
        catch { }
        return "";
    }

    private static EmailResult MapItem(dynamic item, string storeId, string folderName)
    {
        string name = SafeStr(() => item.SenderName);
        string smtp = GetSenderSmtp(item);
        string from =
            name.Length > 0 && smtp.Length > 0 && !string.Equals(name, smtp, StringComparison.OrdinalIgnoreCase)
                ? $"{name} <{smtp}>"
                : name.Length > 0 ? name : smtp;

        bool hasAtt = false;
        try { hasAtt = item.Attachments.Count > 0; } catch { }

        string sid = storeId;
        if (string.IsNullOrEmpty(sid)) sid = SafeStr(() => item.Parent.StoreID);

        // Note: we deliberately do NOT read item.Body here — that forces the full
        // message to load (~ms each) and the preview pane fetches the body on demand.
        return new EmailResult
        {
            EntryId = SafeStr(() => item.EntryID),
            StoreId = sid,
            Subject = SafeStr(() => item.Subject) is { Length: > 0 } s ? s : "(no subject)",
            From = from,
            To = SafeStr(() => item.To),
            ReceivedAt = TryGetDate(item),
            FolderName = folderName,
            BodyPreview = "",
            HasAttachments = hasAtt
        };
    }

    private static DateTime? TryGetDate(dynamic item)
    {
        try { return (DateTime)item.ReceivedTime; } catch { }
        try { return (DateTime)item.SentOn; } catch { }
        return null;
    }

    private static string SafeStr(Func<dynamic> get)
    {
        try { return (string)(get() ?? "") ?? ""; }
        catch { return ""; }
    }

    private static string StripHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";
        var text = Regex.Replace(html, "<(script|style)[^>]*>.*?</\\1>", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        text = Regex.Replace(text, "<[^>]+>", "");
        text = System.Net.WebUtility.HtmlDecode(text);
        return Regex.Replace(text, @"[ \t]{2,}", " ").Trim();
    }

    // Releases a COM RCW; safe to call with null or non-COM values.
    private static void Rel(object? o)
    {
        try { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
        catch { }
    }

    // Builds an Outlook DASL (@SQL) filter to pre-narrow results store-side.
    private static string BuildDasl(SearchOptions o)
    {
        var clauses = new List<string>();

        if (!string.IsNullOrWhiteSpace(o.Subject))
            clauses.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{Esc(o.Subject)}%'");

        if (!string.IsNullOrWhiteSpace(o.From))
        {
            // An "@" means the user picked/typed an address — match the real SMTP property
            // (fast, indexed, and correct for Exchange senders). Otherwise match display name.
            clauses.Add(o.From.Contains('@')
                ? $"\"{PrSenderSmtp}\" LIKE '%{Esc(o.From)}%'"
                : $"\"urn:schemas:httpmail:fromname\" LIKE '%{Esc(o.From)}%'");
        }

        if (o.DateFrom.HasValue)
            clauses.Add($"\"urn:schemas:httpmail:datereceived\" >= '{o.DateFrom.Value:yyyy-MM-dd} 00:00'");

        if (o.DateTo.HasValue)
            clauses.Add($"\"urn:schemas:httpmail:datereceived\" <= '{o.DateTo.Value:yyyy-MM-dd} 23:59'");

        if (!string.IsNullOrWhiteSpace(o.Keyword))
            clauses.Add($"(\"urn:schemas:httpmail:subject\" LIKE '%{Esc(o.Keyword)}%' " +
                        $"OR \"urn:schemas:httpmail:textdescription\" LIKE '%{Esc(o.Keyword)}%')");

        return clauses.Count == 0 ? "" : "@SQL=" + string.Join(" AND ", clauses);
    }

    private static string Esc(string s) => s.Trim().Replace("'", "''");

    public void Dispose() => _sta.Dispose();
}
