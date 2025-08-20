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
            string path = Path.Combine(catFolder, prefix + Sanitize(subject) + ".msg");

            path = GetUniquePath(path);
            mail.SaveAs(path, Outlook.OlSaveAsType.olMSG);
        }

        public static void SaveAttachmentsOnly(Outlook.MailItem mail, string category, string requestNumber)
        {
            string baseFolder = ResolveArchiveRoot();
            string catFolder = EnsureCategoryFolder(baseFolder, category);
            string prefix = string.IsNullOrWhiteSpace(requestNumber) ? "" : Sanitize(requestNumber) + "_";

            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var att = mail.Attachments[i];
                if (IsInline(att)) continue;

                string candidate = Path.Combine(catFolder, prefix + Sanitize(att.FileName));
                candidate = GetUniquePath(candidate);
                att.SaveAsFile(candidate);
            }
        }

        // ---------------- helpers ----------------

        private static string ResolveArchiveRoot()
        {
            string root = null;

            // 1) נעדיף מפתח שמור ב-Settings (אם קיים)
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

            // 2) אחרת – מה-XML ששמור ב-ConfigPath (NB_Archive_Path)
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

            // 3) ברירת מחדל למסמכים\SavedMails
            if (string.IsNullOrWhiteSpace(root))
                root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");

            // ודא נתיב מוחלט ותיקייה קיימת
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
