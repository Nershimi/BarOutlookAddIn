using System;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using BarOutlookAddIn.Helpers; // AddInConfig.Load

namespace BarOutlookAddIn
{
    internal static class SaveEngine
    {
        public static void SaveWholeMail(Outlook.MailItem mail, string category, string requestNumber)
        {
            string baseFolder = ResolveArchiveRoot();
            string catFolder = EnsureCategoryFolder(baseFolder, category);

            string subject = string.IsNullOrWhiteSpace(mail.Subject) ? "NoSubject" : mail.Subject;
            string prefix = string.IsNullOrWhiteSpace(requestNumber) ? "" : Sanitize(requestNumber) + "_";

            // Try numeric allocation via NumeratorService first (ol{N}.msg).
            string path = null;
            try
            {
                int allocated = NumeratorService.GetNextArchiveNumber();
                path = Path.Combine(catFolder, "ol" + allocated + ".msg");
            }
            catch (Exception)
            {
                // fallback to file-system numeric allocator
                try
                {
                    int allocatedFs;
                    path = FileNameAllocator.AllocatePath(catFolder, ".msg", out allocatedFs);
                }
                catch
                {
                    // final fallback: human-readable sanitized name + unique suffix
                    string human = prefix + Sanitize(subject) + ".msg";
                    string initial = Path.Combine(catFolder, human);
                    path = GetUniquePath(initial);
                }
            }

            path = GetUniquePath(path); // ensure full uniqueness just in case
            mail.SaveAs(path, Outlook.OlSaveAsType.olMSG);
        }

        // Modified: when saving attachments we now:
        // - prefer a numeric name from NumeratorService (ol{N}{ext})
        // - fall back to FileNameAllocator and then unique sanitized name
        // - preserve original extension
        // - insert a DB record like other save paths (original filename passed as description)
        // ent/entityName are optional (pass null to skip DB insert)
        public static void SaveAttachmentsOnly(Outlook.MailItem mail, string category, string requestNumber, EntityInfo ent = null, string entityName = null)
        {
            string baseFolder = ResolveArchiveRoot();
            string catFolder = EnsureCategoryFolder(baseFolder, category);
            string prefix = string.IsNullOrWhiteSpace(requestNumber) ? "" : Sanitize(requestNumber) + "_";

            var writer = new ArchiveWriter();

            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var att = mail.Attachments[i];
                if (IsInline(att)) continue;

                string rawName = att.FileName ?? "attachment";
                string ext = Path.GetExtension(rawName);
                if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";

                string candidate;
                int allocatedNumber = 0;
                try
                {
                    // Prefer DB numerator
                    allocatedNumber = NumeratorService.GetNextArchiveNumber();
                    candidate = Path.Combine(catFolder, "ol" + allocatedNumber + ext);
                }
                catch (Exception)
                {
                    // fallback to file-system numeric allocator
                    try
                    {
                        candidate = FileNameAllocator.AllocatePath(catFolder, ext, out allocatedNumber);
                    }
                    catch
                    {
                        // final fallback: use sanitized name + unique suffix
                        string sanitized = prefix + Sanitize(Path.GetFileNameWithoutExtension(rawName));
                        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "attachment";
                        string initial = Path.Combine(catFolder, sanitized + ext);
                        candidate = GetUniquePath(initial);
                    }
                }

                // Ensure uniqueness (defensive)
                candidate = GetUniquePath(candidate);

                try
                {
                    att.SaveAsFile(candidate);

                    // Try DB insert, if caller provided entity info or name
                    try
                    {
                        // For compatibility, pass entity object if available, else string name (may be empty)
                        if (ent != null)
                            writer.TryInsertRecord(ent, requestNumber, candidate, rawName);
                        else
                            writer.TryInsertRecord(entityName ?? string.Empty, requestNumber, candidate, rawName);
                    }
                    catch (Exception exDb)
                    {
                        DevDiag.Log("SaveAttachmentsOnly: DB insert EX: " + exDb.Message);
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    DevDiag.Log("SaveAttachmentsOnly: att.SaveAsFile COMEX 0x" + comEx.HResult.ToString("X") + ": " + comEx.Message);
                }
                catch (Exception ex)
                {
                    DevDiag.Log("SaveAttachmentsOnly: att.SaveAsFile EX: " + ex.Message);
                }
            }
        }

        // ---------------- helpers ----------------

        private static string ResolveArchiveRoot()
        {
            string root = null;

            // 1) prefer Settings["SaveBaseFolder"]
            try
            {
                var props = Properties.Settings.Default.Properties;
                if (props != null && props["SaveBaseFolder"] != null)
                {
                    var val = Properties.Settings.Default["SaveBaseFolder"] as string;
                    if (!string.IsNullOrWhiteSpace(val)) root = val;
                }
            }
            catch { }

            // 2) else from XML at ConfigPath (NB_Archive_Path)
            if (string.IsNullOrWhiteSpace(root))
            {
                try
                {
                    string cfgPath = Properties.Settings.Default.ConfigPath;
                    if (!string.IsNullOrWhiteSpace(cfgPath) && File.Exists(cfgPath))
                    {
                        var cfg = AddInConfig.Load(cfgPath);
                        if (cfg != null && !string.IsNullOrWhiteSpace(cfg.ArchivePath))
                            root = cfg.ArchivePath;
                    }
                }
                catch { }
            }

            // 3) default Documents\SavedMails
            if (string.IsNullOrWhiteSpace(root))
                root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");

            // ensure absolute path and create
            if (!Path.IsPathRooted(root))
                root = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), root));

            try { Directory.CreateDirectory(root); }
            catch
            {
                string fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");
                Directory.CreateDirectory(fallback);
                root = fallback;
            }

            return root;
        }

        private static string EnsureCategoryFolder(string baseFolder, string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                Directory.CreateDirectory(baseFolder);
                return baseFolder;
            }

            string full = Path.Combine(baseFolder, Sanitize(category));
            Directory.CreateDirectory(full);
            return full;
        }


        private static string GetUniquePath(string path)
        {
            if (!File.Exists(path)) return path;
            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            for (int i = 1; i < 1000; i++)
            {
                string cand = Path.Combine(dir, name + "_" + i + ext);
                if (!File.Exists(cand)) return cand;
            }
            return Path.Combine(dir, name + "_" + DateTime.Now.Ticks + ext);
        }

        private static string Sanitize(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Unnamed";
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }

        private static bool IsInline(Outlook.Attachment att)
        {
            try
            {
                return att != null && (att.Type == Outlook.OlAttachmentType.olOLE ||
                        att.FileName.EndsWith(".htm", StringComparison.OrdinalIgnoreCase));
            }
            catch { return false; }
        }
    }
}
