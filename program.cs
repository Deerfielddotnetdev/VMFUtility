// MailFlowEmlExporter, Version=4.2
// MailFlow Super Utility - EML export, Soft/Hard Purge, Totals + Simple Unlock

using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MailFlowSuperUtility
{
    internal static class Program
    {
        private const int CommandTimeoutSeconds = 300;
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailFlowUtility.log");

        [STAThread]
        private static void Main(string[] args)
        {
            // --- Simple unlock: show the form; proceed only on OK ---
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var form = new UnlockForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                    return;
            }

            Log("=== Starting MailFlow Super Utility ===");
            using (var conn = BuildAndOpenConnection(args))
            {
                if (conn == null) return;

                while (true)
                {
                    Console.WriteLine();
                    Console.WriteLine("Select an option:");
                    Console.WriteLine("1 - Export Inbound Messages to EML");
                    Console.WriteLine("2 - Export Outbound Messages to EML");
                    Console.WriteLine("3 - Export Both (Inbound + Outbound) to EML");
                    Console.WriteLine("4 - Mark Tickets for Purge (Soft Delete)");
                    Console.WriteLine("5 - Finalize and Hard Purge Marked Tickets by Date Range");
                    Console.WriteLine("6 - MailFlow Totals (Agents & Ticket counts)");
                    Console.WriteLine("0 - Exit");
                    Console.Write("Choice: ");
                    var choice = Console.ReadLine();

                    try
                    {
                        switch (choice)
                        {
                            case "1":
                            {
                                var start = ReadDate("Enter START date");
                                var end   = ReadDate("Enter END date");
                                ExportMessages(conn, start, end, new[] { "Inbound" });
                                break;
                            }
                            case "2":
                            {
                                var start = ReadDate("Enter START date");
                                var end   = ReadDate("Enter END date");
                                ExportMessages(conn, start, end, new[] { "Outbound" });
                                break;
                            }
                            case "3":
                            {
                                var start = ReadDate("Enter START date");
                                var end   = ReadDate("Enter END date");
                                ExportMessages(conn, start, end, new[] { "Inbound", "Outbound" });
                                break;
                            }
                            case "4":
                                DoSoftDelete(conn);
                                break;
                            case "5":
                                DoHardPurge(conn);
                                break;
                            case "6":
                                ShowTotals(conn);
                                break;
                            case "0":
                                Log("Exiting.");
                                return;
                            default:
                                Console.WriteLine("Unknown option. Try again.");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("ERROR: " + ex);
                    }
                }
            }
        }

        // ------------ Logging ------------
        private static void Log(string message)
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            Console.WriteLine(line);
            File.AppendAllText(LogPath, line + Environment.NewLine);
        }

        // ------------ Input helpers ------------
        private static DateTime ReadDate(string prompt)
        {
            while (true)
            {
                Console.Write($"{prompt} (YYYY-MM-DD): ");
                var text = Console.ReadLine()?.Trim();
                if (DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    return dt;
                Console.WriteLine("Invalid date. Please enter in YYYY-MM-DD format.");
            }
        }

        private static int ReadIntOrDefault(string prompt, int @default)
        {
            Console.Write($"{prompt} (default {@default}): ");
            var txt = Console.ReadLine();
            return int.TryParse(txt, out var val) ? val : @default;
        }

        private static string ReadPassword()
        {
            var s = "";
            ConsoleKeyInfo k;
            do
            {
                k = Console.ReadKey(intercept: true);
                if (k.Key != ConsoleKey.Backspace && k.Key != ConsoleKey.Enter)
                {
                    s += k.KeyChar;
                    Console.Write("*");
                }
                else if (k.Key == ConsoleKey.Backspace && s.Length > 0)
                {
                    s = s.Substring(0, s.Length - 1);
                    Console.Write("\b \b");
                }
            } while (k.Key != ConsoleKey.Enter);
            Console.WriteLine();
            return s;
        }

        // ------------ DB bootstrap ------------
        private static SqlConnection BuildAndOpenConnection(string[] args)
        {
            string server = null, db = null, auth = null, user = null, pass = null;

            foreach (var arg in args)
            {
                var parts = arg.Split('=');
                if (parts.Length != 2) continue;
                var key = parts[0].Trim().ToLowerInvariant();
                var val = parts[1].Trim();
                switch (key)
                {
                    case "--server": server = val; break;
                    case "--db": db = val; break;
                    case "--auth": auth = val; break;
                    case "--user": user = val; break;
                    case "--pass": pass = val; break;
                }
            }

            while (true)
            {
                try
                {
                    if (string.IsNullOrEmpty(server))
                    {
                        Console.Write("Enter SQL Server name (e.g., localhost\\SQLEXPRESS): ");
                        server = Console.ReadLine();
                    }
                    if (string.IsNullOrEmpty(db))
                    {
                        Console.Write("Enter Database name: ");
                        db = Console.ReadLine();
                    }
                    if (string.IsNullOrEmpty(auth))
                    {
                        Console.WriteLine("Select Authentication Mode:");
                        Console.WriteLine("1 - Windows Integrated Security");
                        Console.WriteLine("2 - SQL Server Authentication");
                        Console.Write("Choice: ");
                        auth = Console.ReadLine();
                    }

                    var csb = new SqlConnectionStringBuilder
                    {
                        DataSource = server,
                        InitialCatalog = db,
                        MultipleActiveResultSets = true
                    };

                    if (auth == "1" || string.Equals(auth, "windows", StringComparison.OrdinalIgnoreCase))
                    {
                        csb.IntegratedSecurity = true;
                    }
                    else if (auth == "2" || string.Equals(auth, "sql", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(user))
                        {
                            Console.Write("SQL Username: ");
                            user = Console.ReadLine();
                        }
                        if (string.IsNullOrEmpty(pass))
                        {
                            Console.Write("SQL Password: ");
                            pass = ReadPassword();
                        }
                        csb.UserID = user;
                        csb.Password = pass;
                        csb.IntegratedSecurity = false;
                    }
                    else
                    {
                        Log("Invalid authentication option. Exiting.");
                        return null;
                    }

                    Log("Connecting to database...");
                    var conn = new SqlConnection(csb.ConnectionString);
                    conn.Open();
                    Log("Successfully logged in.");
                    return conn;
                }
                catch (Exception ex)
                {
                    Log("Connection failed: " + ex.Message);
                    Console.Write("Would you like to retry? (y/n): ");
                    if (!string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
                        return null;
                }
            }
        }

        // ------------ DB helpers ------------
        private static int ExecNonQuery(SqlConnection conn, string sql, Action<SqlCommand> bind, SqlTransaction tx = null)
        {
            using (var cmd = (tx == null) ? new SqlCommand(sql, conn) : new SqlCommand(sql, conn, tx))
            {
                cmd.CommandTimeout = CommandTimeoutSeconds;
                bind?.Invoke(cmd);
                return cmd.ExecuteNonQuery();
            }
        }

        private static T ExecScalar<T>(SqlConnection conn, string sql, Action<SqlCommand> bind, SqlTransaction tx = null)
        {
            using (var cmd = (tx == null) ? new SqlCommand(sql, conn) : new SqlCommand(sql, conn, tx))
            {
                cmd.CommandTimeout = CommandTimeoutSeconds;
                bind?.Invoke(cmd);
                object result = cmd.ExecuteScalar();
                return (T)Convert.ChangeType(result, typeof(T), CultureInfo.InvariantCulture);
            }
        }

        // ------------ Feature: EML export ------------
        private static void ExportMessages(SqlConnection conn, DateTime start, DateTime end, string[] kinds, string outputBase = "exports")
        {
            Directory.CreateDirectory(outputBase);

            foreach (var kind in kinds)
            {
                string msgTable = kind + "Messages";
                string attJoinTable = kind + "MessageAttachments";
                string idCol = kind + "MessageID";

                string sql = $@"
SELECT
    M.{idCol},
    M.EmailFrom,
    M.EmailPrimaryTo,
    M.EmailTo,
    M.EmailCc,
    M.EmailReplyTo,
    M.EmailDateTime,
    M.Subject,
    M.Body,
    M.MediaSubType,
    TBox.Name AS TicketBoxName
FROM {msgTable} M
LEFT JOIN Tickets TK   ON M.TicketID = TK.TicketID
LEFT JOIN TicketBoxes TBox ON TK.TicketBoxID = TBox.TicketBoxID
WHERE M.EmailDateTime BETWEEN @StartDate AND @EndDate
ORDER BY M.EmailDateTime";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.CommandTimeout = CommandTimeoutSeconds;
                    cmd.Parameters.AddWithValue("@StartDate", start);
                    cmd.Parameters.AddWithValue("@EndDate", end);

                    using (var rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            int msgId = rdr.GetInt32(0);
                            string emailFrom = rdr["EmailFrom"]?.ToString() ?? "";
                            string emailTo = rdr["EmailTo"]?.ToString() ?? "";
                            string emailPrimaryTo = rdr["EmailPrimaryTo"]?.ToString() ?? "";
                            if (!string.IsNullOrWhiteSpace(emailPrimaryTo))
                                emailTo = string.IsNullOrWhiteSpace(emailTo) ? emailPrimaryTo : (emailPrimaryTo + ";" + emailTo);
                            string emailCc = rdr["EmailCc"]?.ToString() ?? "";
                            string subject = rdr["Subject"]?.ToString() ?? "";
                            string body = rdr["Body"]?.ToString() ?? "";
                            string mediaSubType = rdr["MediaSubType"]?.ToString()?.ToLowerInvariant() ?? "";
                            DateTime sent = rdr.GetDateTime(rdr.GetOrdinal("EmailDateTime"));
                            string ticketBox = rdr["TicketBoxName"]?.ToString() ?? "UnknownBox";

                            subject = new string(subject.Where(c => !char.IsControl(c)).ToArray()).Trim();
                            if (string.IsNullOrWhiteSpace(subject)) subject = "No Subject";
                            string safeSubject = Regex.Replace(subject, "[\\/:*?\"<>|]", "_");
                            if (safeSubject.Length > 120) safeSubject = safeSubject.Substring(0, 120);

                            string safeBox = Regex.Replace(ticketBox, "[\\/:*?\"<>|]", "_");
                            string outDir = Path.Combine(outputBase, safeBox, kind);
                            Directory.CreateDirectory(outDir);
                            string outPath = Path.Combine(outDir, msgId + "_" + safeSubject + ".eml");

                            using (var mail = new MailMessage())
                            {
                                try { mail.From = new MailAddress(emailFrom); }
                                catch { mail.From = new MailAddress("noreply@example.com"); }

                                AddAddresses(mail.To, emailTo);
                                AddAddresses(mail.CC, emailCc);

                                mail.Subject = subject;
                                mail.Body = body;
                                mail.IsBodyHtml = (mediaSubType == "html");
                                mail.Headers["Date"] = sent.ToString("r");

                                using (var cmdAtt = new SqlCommand(
                                    "SELECT A.AttachmentLocation, A.FileName " +
                                    "FROM " + attJoinTable + " MA " +
                                    "JOIN Attachments A ON MA.AttachmentID = A.AttachmentID " +
                                    "WHERE MA." + idCol + " = @MsgID", conn))
                                {
                                    cmdAtt.CommandTimeout = CommandTimeoutSeconds;
                                    cmdAtt.Parameters.AddWithValue("@MsgID", msgId);
                                    using (var r2 = cmdAtt.ExecuteReader())
                                    {
                                        while (r2.Read())
                                        {
                                            var path = r2["AttachmentLocation"]?.ToString();
                                            var name = r2["FileName"]?.ToString();
                                            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                                            {
                                                var att = new Attachment(path);
                                                if (!string.IsNullOrWhiteSpace(name)) att.Name = name;
                                                mail.Attachments.Add(att);
                                            }
                                        }
                                    }
                                }

                                // Unique temp pickup dir per message (avoid races)
                                string pickDir = Path.Combine(Path.GetTempPath(), "MailFlowEmlExporter", Guid.NewGuid().ToString("N"));
                                Directory.CreateDirectory(pickDir);

                                using (var smtp = new SmtpClient())
                                {
                                    smtp.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                                    smtp.PickupDirectoryLocation = pickDir;
                                    smtp.Send(mail);
                                }

                                var emlFile = Directory.EnumerateFiles(pickDir, "*.eml").FirstOrDefault();
                                if (emlFile == null)
                                    throw new InvalidOperationException("No EML file was created in pickup directory.");

                                Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? outDir);
                                if (File.Exists(outPath)) File.Delete(outPath); // Framework-safe overwrite
                                File.Move(emlFile, outPath);

                                try { Directory.Delete(pickDir, true); } catch { /* ignore */ }

                                Log("Exported " + kind + " message ID " + msgId + " to " + outPath);
                            }
                        }
                    }
                }
            }

            Log("Export complete.");
        }

        private static void AddAddresses(MailAddressCollection col, string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            var parts = raw.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var addr = p.Trim();
                if (addr.Length == 0) continue;
                try { col.Add(addr); } catch { /* skip invalid */ }
            }
        }

        // ------------ Feature: Soft delete (clarified) ------------
        private static void DoSoftDelete(SqlConnection conn)
        {
            var start = ReadDate("Enter START date");
            var end   = ReadDate("Enter END date");

            // Show a quick summary of in-range states to avoid mistakes
            PrintStateSummary(conn, start, end);

            Console.WriteLine("Common TicketStateIDs: 1=Closed, 2=Open, 3=On-Hold, 6=Marked for Deletion.");
            int defaultStateId = 1; // keep default as Closed
            int stateId = ReadIntOrDefault($"Enter TicketStateID to mark as deleted (default {defaultStateId} = {GetStateLabel(defaultStateId)} )", defaultStateId);

            // Preview how many will be affected
            const string previewSql = @"
SELECT COUNT(*) 
FROM Tickets 
WHERE IsDeleted = 0
  AND TicketStateID = @StateID
  AND DateCreated BETWEEN @StartDate AND @EndDate;";
            int previewCount = ExecScalar<int>(conn, previewSql, cmd =>
            {
                cmd.Parameters.AddWithValue("@StateID", stateId);
                cmd.Parameters.AddWithValue("@StartDate", start);
                cmd.Parameters.AddWithValue("@EndDate", end);
            });

            Console.Write($"This will mark {previewCount} tickets as deleted. Proceed? (y/n): ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
            {
                Log("Soft delete cancelled.");
                return;
            }

            var deletedBy = ReadIntOrDefault("Enter AgentID for DeletedBy", 1);

            const string updateSql = @"
UPDATE Tickets
   SET IsDeleted = 1,
       DeletedBy = @DeletedBy,
       DeletedTime = GETDATE()
 WHERE TicketStateID = @StateID
   AND DateCreated BETWEEN @StartDate AND @EndDate
   AND IsDeleted = 0;";

            Log($"Marking IsDeleted=1 from {start:yyyy-MM-dd} to {end:yyyy-MM-dd}, TicketStateID={stateId} ({GetStateLabel(stateId)}), DeletedBy={deletedBy}...");
            int affected = ExecNonQuery(conn, updateSql, cmd =>
            {
                cmd.Parameters.AddWithValue("@DeletedBy", deletedBy);
                cmd.Parameters.AddWithValue("@StateID", stateId);
                cmd.Parameters.AddWithValue("@StartDate", start);
                cmd.Parameters.AddWithValue("@EndDate", end);
            });

            Log($"Total tickets marked for purge: {affected}");
        }

        private static string GetStateLabel(int id) => id switch
        {
            1 => "Closed",
            2 => "Open",
            3 => "On-Hold",
            6 => "Marked for Deletion",
            _ => $"State {id}"
        };

        private static void PrintStateSummary(SqlConnection conn, DateTime start, DateTime end)
        {
            const string sql = @"
IF OBJECT_ID('TicketStates','U') IS NOT NULL
BEGIN
    SELECT t.TicketStateID, s.Name, COUNT(*) AS Cnt
    FROM Tickets t
    JOIN TicketStates s ON s.TicketStateID = t.TicketStateID
    WHERE t.DateCreated BETWEEN @Start AND @End
      AND t.IsDeleted = 0
    GROUP BY t.TicketStateID, s.Name
    ORDER BY t.TicketStateID
END
ELSE
BEGIN
    SELECT t.TicketStateID, CAST(NULL AS NVARCHAR(200)) AS Name, COUNT(*) AS Cnt
    FROM Tickets t
    WHERE t.DateCreated BETWEEN @Start AND @End
      AND t.IsDeleted = 0
    GROUP BY t.TicketStateID
    ORDER BY t.TicketStateID
END";

            using var cmd = new SqlCommand(sql, conn) { CommandTimeout = CommandTimeoutSeconds };
            cmd.Parameters.AddWithValue("@Start", start);
            cmd.Parameters.AddWithValue("@End", end);

            Console.WriteLine();
            Console.WriteLine("Ticket states in the selected date range (IsDeleted = 0):");

            using var r = cmd.ExecuteReader();
            bool any = false;
            while (r.Read())
            {
                any = true;
                int id = r.GetInt32(0);
                string name = r.IsDBNull(1) ? GetStateLabel(id) : r.GetString(1);
                int cnt = r.GetInt32(2);
                Console.WriteLine($"  {id} = {name,-20} : {cnt,6}");
            }

            if (!any)
                Console.WriteLine("  (No tickets found in this date range.)");

            Console.WriteLine();
        }

        // ------------ Feature: Hard purge ------------
        private static void DoHardPurge(SqlConnection conn)
        {
            var start = ReadDate("Enter START date for hard purge");
            var end = ReadDate("Enter END date for hard purge");

            Console.Write("Are you sure you want to permanently delete marked tickets in this range? (y/n): ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
                return;

            Log("Starting hard purge...");

            using (var tx = conn.BeginTransaction())
            {
                ExecNonQuery(conn, "SET XACT_ABORT ON;", _ => { }, tx);

                string ticketScope = @"
WITH ToDelete AS (
    SELECT TicketID
      FROM Tickets
     WHERE IsDeleted = 1
       AND DateCreated BETWEEN @StartDate AND @EndDate
)";

                int toDeleteCount = ExecScalar<int>(conn, ticketScope + " SELECT COUNT(*) FROM ToDelete;", cmd =>
                {
                    cmd.Parameters.AddWithValue("@StartDate", start);
                    cmd.Parameters.AddWithValue("@EndDate", end);
                }, tx);

                Log("Tickets to hard delete: " + toDeleteCount);

                Action<string, string> Delete = (label, sql) =>
                {
                    int rows = ExecNonQuery(conn, ticketScope + sql, cmd =>
                    {
                        cmd.Parameters.AddWithValue("@StartDate", start);
                        cmd.Parameters.AddWithValue("@EndDate", end);
                    }, tx);
                    Log(label + ": " + rows + " rows deleted");
                };

                try
                {
                    // Children first, parents last
                    Delete("OutboundMessageAttachments",
                        "DELETE FROM OutboundMessageAttachments WHERE OutboundMessageID IN (SELECT OutboundMessageID FROM OutboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("InboundMessageAttachments",
                        "DELETE FROM InboundMessageAttachments WHERE InboundMessageID IN (SELECT InboundMessageID FROM InboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("NoteAttachments",
                        "DELETE FROM NoteAttachments WHERE NoteID IN (SELECT TicketNoteID FROM TicketNotes WHERE TicketID IN (SELECT TicketID FROM ToDelete) AND NoteTypeID = 1)");
                    Delete("InboundMessageRead",
                        "DELETE FROM InboundMessageRead WHERE InboundMessageID IN (SELECT InboundMessageID FROM InboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("InboundMessageQueue",
                        "DELETE FROM InboundMessageQueue WHERE InboundMessageID IN (SELECT InboundMessageID FROM InboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("SRKeywordResults",
                        "DELETE FROM SRKeywordResults WHERE InboundMessageID IN (SELECT InboundMessageID FROM InboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("TicketNotesRead",
                        "DELETE FROM TicketNotesRead WHERE TicketNoteID IN (SELECT TicketNoteID FROM TicketNotes WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("InboundMessages",
                        "DELETE FROM InboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("OutboundMessageContacts",
                        "DELETE FROM OutboundMessageContacts WHERE OutboundMessageID IN (SELECT OutboundMessageID FROM OutboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("OutboundMessageQueue",
                        "DELETE FROM OutboundMessageQueue WHERE OutboundMessageID IN (SELECT OutboundMessageID FROM OutboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete))");
                    Delete("OutboundMessages",
                        "DELETE FROM OutboundMessages WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("TicketNotes",
                        "DELETE FROM TicketNotes WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("TicketFieldsTicket",
                        "DELETE FROM TicketFieldsTicket WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("TicketLinksTicket",
                        "DELETE FROM TicketLinksTicket WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("TicketContacts",
                        "DELETE FROM TicketContacts WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("TicketHistory",
                        "DELETE FROM TicketHistory WHERE TicketID IN (SELECT TicketID FROM ToDelete)");
                    Delete("Tickets",
                        "DELETE FROM Tickets WHERE TicketID IN (SELECT TicketID FROM ToDelete)");

                    tx.Commit();
                    Log("Hard purge complete.");
                }
                catch (Exception ex)
                {
                    try { tx.Rollback(); } catch { }
                    throw new Exception("Hard purge failed and was rolled back.", ex);
                }
            }
        }

        // ------------ Feature: Totals ------------
        private static void ShowTotals(SqlConnection conn)
        {
            try
            {
                Log("=== MailFlow Totals ===");

                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM Agents WHERE IsEnabled = 1", conn))
                {
                    cmd.CommandTimeout = CommandTimeoutSeconds;
                    Log("Total enabled agents: " + (int)cmd.ExecuteScalar());
                }

                const string byStatusSql = @"
SELECT TicketStateID, COUNT(*) AS Count
FROM Tickets
WHERE TicketStateID IN (1,2,3,6)
GROUP BY TicketStateID
ORDER BY TicketStateID";

                using (var cmd2 = new SqlCommand(byStatusSql, conn))
                {
                    cmd2.CommandTimeout = CommandTimeoutSeconds;
                    using (var rdr = cmd2.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            int state = rdr.GetInt32(0);
                            int count = rdr.GetInt32(1);
                            string label = (state == 1) ? "Closed"
                                : (state == 2) ? "Open"
                                : (state == 3) ? "On-Hold"
                                : (state == 6) ? "Marked for Deletion"
                                : "Unknown State (" + state + ")";
                            Log("Tickets (" + label + "): " + count);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Error while retrieving MailFlow totals: " + ex.Message);
            }
        }
    }
}
