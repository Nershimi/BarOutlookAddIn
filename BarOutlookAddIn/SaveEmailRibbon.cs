using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using BarOutlookAddIn.Helpers;

namespace BarOutlookAddIn
{
    public partial class SaveEmailRibbon
    {
        private void SaveEmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            DevDiag.Log("Designer: SaveEmailRibbon_Load");
        }

        private void btnSaveSelectedEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DevDiag.Log("Ribbon: click");

                Outlook.Application app = Globals.ThisAddIn.Application;

                // Try to get current MailItem
                Outlook.MailItem mailItem = null;
                var insp = app.ActiveInspector();
                if (insp != null) mailItem = insp.CurrentItem as Outlook.MailItem;

                if (mailItem == null)
                {
                    var selection = app.ActiveExplorer()?.Selection;
                    if (selection != null && selection.Count > 0 && selection[1] is Outlook.MailItem mi)
                        mailItem = mi;
                }

                DevDiag.Log("Ribbon: mailItem? " + (mailItem != null));
                if (mailItem == null) { MessageBox.Show("לא נבחר מייל."); return; }

                using (var dialog = new SaveEmailDialog())
                {
                    var dr = dialog.ShowDialog();
                    DevDiag.Log("Ribbon: dialog result = " + dr);
                    if (dr != DialogResult.OK) return;

                    // Entity selection (נבדוק מה באמת נבחר)
                    var ent = dialog.SelectedEntityInfo;
                    string entityName = dialog.SelectedEntityName ?? string.Empty;
                    DevDiag.Log("Ribbon: entity after dialog -> " +
                        (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>")
                        + " | entityNameStr='" + entityName + "'");

                    // Resolve root/category
                    string baseFolder = ResolveArchiveRoot();
                    string category = dialog.SelectedCategory ?? string.Empty;
                    string requestNumber = dialog.RequestNumber ?? string.Empty;
                    string categoryFolder = EnsureCategoryFolder(baseFolder, category);

                    DevDiag.Log("Ribbon: baseFolder=" + baseFolder
                        + " | category='" + category + "'"
                        + " | categoryFolder=" + categoryFolder
                        + " | req='" + requestNumber + "'");

                    // וידוא הרשאות כתיבה פיזיות
                    if (!EnsureWritableFolder(categoryFolder, out string ensureErr))
                    {
                        DevDiag.Log("Ribbon: EnsureWritableFolder FAIL: " + ensureErr);
                        MessageBox.Show("אין אפשרות לכתוב לתיקייה: " + categoryFolder + "\r\n" + ensureErr,
                            "שגיאת גישה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    DevDiag.Log("Ribbon: EnsureWritableFolder OK");

                    var writer = new ArchiveWriter();
                    DevDiag.Log("Ribbon: ArchiveWriter created");

                    DevDiag.Log("Ribbon: SelectedOption=" + dialog.SelectedOption);

                    // --- 1) Save whole email (.msg) named by the Subject ---
                    // --- 1) Save whole email (.msg) named by the Subject or by custom name ---
                    if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveEmail)
                    {
                        // אם המשתמש סימן "שם קובץ מותאם" – נשתמש בו, אחרת בנושא המייל
                        string fileBase = dialog.UseCustomFileName
                            ? CleanFileName(dialog.CustomFileName)
                            : CleanFileName(mailItem.Subject ?? "NoSubject");
                        if (string.IsNullOrWhiteSpace(fileBase)) fileBase = "NoSubject";

                        string initialPath = Path.Combine(categoryFolder, fileBase + ".msg");
                        string fullPath = GetUniquePath(initialPath);

                        DevDiag.Log($"Ribbon: saving MSG. base='{fileBase}', path='{fullPath}'");

                        mailItem.SaveAs(fullPath, Outlook.OlSaveAsType.olMSG);
                        DevDiag.Log("Ribbon: saved msg exists? " + File.Exists(fullPath));

                        try
                        {
                            bool ok = (ent != null)
                                ? writer.TryInsertRecord(ent, requestNumber, fullPath, Path.GetFileName(fullPath))
                                : writer.TryInsertRecord(entityName, requestNumber, fullPath, Path.GetFileName(fullPath));
                            DevDiag.Log("Ribbon: DB insert after MSG ok? " + ok);
                        }
                        catch (Exception exDb)
                        {
                            DevDiag.Log("Ribbon: DB insert after MSG EX: " + exDb.Message);
                        }

                        MessageBox.Show("המייל נשמר:\n" + fullPath, "שמירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    // --- 2) Save only attachments, preserving original file names ---
                    else if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveAttachmentsOnly
      || dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveAttachments)
                    {
                        if (mailItem.Attachments.Count == 0)
                        {
                            MessageBox.Show("אין קבצים מצורפים במייל.");
                            return;
                        }

                        var validAttachments = new List<Outlook.Attachment>();
                        foreach (Outlook.Attachment att in mailItem.Attachments)
                            if (!IsInlineImage(att)) validAttachments.Add(att);

                        DevDiag.Log("Ribbon: attachments total=" + mailItem.Attachments.Count + ", valid=" + validAttachments.Count);

                        if (validAttachments.Count == 0)
                        {
                            MessageBox.Show("לא נמצאו קבצים תקינים לשמירה.");
                            return;
                        }

                        var names = validAttachments.Select(a => a.FileName).ToList();
                        using (var selectionDialog = new AttachmentSelectionDialog(names))
                        {
                            if (selectionDialog.ShowDialog() != DialogResult.OK)
                                return;

                            var selectedSet = new HashSet<string>(selectionDialog.SelectedAttachments, StringComparer.OrdinalIgnoreCase);
                            int selCount = selectedSet.Count;
                            int counter = 0;
                            int saved = 0;

                            foreach (var att in validAttachments)
                            {
                                if (!selectedSet.Contains(att.FileName))
                                    continue;

                                counter++;

                                string rawName = att.FileName ?? "";
                                string ext = Path.GetExtension(rawName);
                                if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";

                                // בסיס השם: מותאם/ממוספר או לפי שם המצורף
                                string baseForThis =
                                    dialog.UseCustomFileName
                                        ? (selCount == 1
                                            ? CleanFileName(dialog.CustomFileName)
                                            : CleanFileName(dialog.CustomFileName) + "_" + counter)
                                        : CleanFileName(Path.GetFileNameWithoutExtension(rawName) ?? "attachment");

                                if (string.IsNullOrWhiteSpace(baseForThis)) baseForThis = "attachment";

                                string initialPath = Path.Combine(categoryFolder, baseForThis + ext);
                                string filePath = GetUniquePath(initialPath);

                                DevDiag.Log($"Ribbon: saving ATT. raw='{rawName}', base='{baseForThis}', ext='{ext}', path='{filePath}'");

                                try
                                {
                                    att.SaveAsFile(filePath);
                                    DevDiag.Log("Ribbon: saved att exists? " + File.Exists(filePath));
                                    saved++;

                                    try
                                    {
                                        bool ok = (ent != null)
                                          ? writer.TryInsertRecord(ent, requestNumber, filePath, Path.GetFileName(filePath))
                                          : writer.TryInsertRecord(entityName, requestNumber, filePath, Path.GetFileName(filePath));
                                        DevDiag.Log("Ribbon: DB insert after ATT ok? " + ok);
                                    }
                                    catch (Exception exDbA)
                                    {
                                        DevDiag.Log("Ribbon: DB insert after ATT EX: " + exDbA.Message);
                                    }
                                }
                                catch (System.Runtime.InteropServices.COMException comEx)
                                {
                                    DevDiag.Log("Ribbon: att.SaveAsFile COMEX 0x" + comEx.HResult.ToString("X") + ": " + comEx.Message);
                                }
                                catch (Exception exAtt)
                                {
                                    DevDiag.Log("Ribbon: att.SaveAsFile EX: " + exAtt.Message);
                                }
                            }

                            MessageBox.Show(
                                saved > 0 ? "הקבצים שנבחרו נשמרו בהצלחה." : "לא נבחרו קבצים לשמירה.",
                                "שמירת קבצים", MessageBoxButtons.OK,
                                saved > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                        }
                    }

                }
            }
            catch (COMException comRoot)
            {
                DevDiag.Log("Ribbon: ROOT COMEX 0x" + comRoot.HResult.ToString("X") + ": " + comRoot.Message);
                MessageBox.Show("שגיאת COM:\r\n" + comRoot.Message);
            }
            catch (Exception ex)
            {
                DevDiag.Log("Ribbon: EX " + ex.Message);
                MessageBox.Show("שגיאה: " + ex.Message);
            }
        }


        // בודק הרשאות כתיבה ע"י יצירת קובץ זמני ומחיקתו
        private bool EnsureWritableFolder(string folder, out string error)
        {
            error = null;
            try
            {
                Directory.CreateDirectory(folder);
                string probe = Path.Combine(folder, "~write_probe_" + Guid.NewGuid().ToString("N") + ".tmp");
                using (var fs = File.Create(probe)) { }
                File.Delete(probe);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // בתוך המחלקה SaveEmailRibbon
        private string GetUniquePath(string initialPath)
        {
            if (string.IsNullOrWhiteSpace(initialPath))
                throw new ArgumentException("initialPath is null/empty");

            string dir = Path.GetDirectoryName(initialPath);
            string name = Path.GetFileNameWithoutExtension(initialPath);
            string ext = Path.GetExtension(initialPath);

            if (string.IsNullOrWhiteSpace(dir))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (ext == null) ext = string.Empty;

            try { Directory.CreateDirectory(dir); } catch { /* ignore */ }

            // 1) try original name
            string candidate = Path.Combine(dir, name + ext);
            if (!File.Exists(candidate)) return candidate;

            // 2) try suffixes _1, _2, ...
            for (int i = 1; i < 1000; i++)
            {
                candidate = Path.Combine(dir, name + "_" + i + ext);
                if (!File.Exists(candidate)) return candidate;
            }

            // 3) fallback with timestamp
            candidate = Path.Combine(dir, name + "_" +
                       DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ext);
            return candidate;
        }


        // שמירת MSG עם לכידת COMException + פירוט
        private bool TrySaveMsgWithDiag(Outlook.MailItem mail, string path, out string error)
        {
            error = null;
            try
            {
                mail.SaveAs(path, Outlook.OlSaveAsType.olMSG);
                return true;
            }
            catch (COMException comEx)
            {
                error = "COM 0x" + comEx.HResult.ToString("X") + ": " + comEx.Message;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // ----------------- Helpers -----------------

        // 1) Root folder: Settings["SaveBaseFolder"] → XML NB_Archive_Path → Documents\SavedMails
        private string ResolveArchiveRoot()
        {
            string root = null;

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

            if (string.IsNullOrWhiteSpace(root))
                root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");

            if (!Path.IsPathRooted(root))
            {
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                root = Path.GetFullPath(Path.Combine(docs, root));
            }

            try { Directory.CreateDirectory(root); }
            catch
            {
                string fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");
                Directory.CreateDirectory(fallback);
                root = fallback;
            }

            DevDiag.Log("Ribbon: ResolveArchiveRoot -> " + root);
            return root;
        }

        // 2) Create category subfolder (or use base when empty)
        private string EnsureCategoryFolder(string baseFolder, string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                Directory.CreateDirectory(baseFolder);
                return baseFolder;
            }

            string full = Path.Combine(baseFolder, CleanFileName(category));
            Directory.CreateDirectory(full);
            return full;
        }

        // 3) Sanitize file/folder name
        private static string CleanFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Unnamed";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }

        // 4) Detect inline/signature attachments
        // using Outlook = Microsoft.Office.Interop.Outlook;
        private bool IsInlineImage(Outlook.Attachment att)
        {
            try
            {
                if (att == null) return false;

                string name = att.FileName ?? "";
                string ext = System.IO.Path.GetExtension(name).ToLowerInvariant();
                string cid = GetContentId(att); // null אם אין

                // לוג אבחוני
                DevDiag.Log($"InlineCheck: name='{name}', ext='{ext}', type={att.Type}, cid='{cid ?? "<null>"}', size={att.Size}");

                // HTML חלקי גוף
                if (ext == ".htm" || ext == ".html") return true;

                // תמונות—נטפל בהן כ-inline רק אם יש ContentId (כלומר משובצות בגוף HTML)
                if (ext == ".gif" || ext == ".jpg" || ext == ".jpeg" || ext == ".png")
                    return !string.IsNullOrEmpty(cid);

                // לא חוסמים יותר olOLE באופן גורף, כי זה עלול לכלול DWG/PDF/Excel מוטמעים.
                // אם יש ContentId אבל זו *לא* תמונת־web—נעדיף לשמור (יש מיילים עם CID גם לקבצים רגילים).
                return false;
            }
            catch { return false; }
        }

        private string GetContentId(Outlook.Attachment att)
        {
            try
            {
                // PR_ATTACH_CONTENT_ID (0x3712, PT_TSTRING)
                return att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E") as string;
            }
            catch { return null; }
        }

    }
}
