using BarOutlookAddIn.Helpers;
using stdole;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;        

namespace BarOutlookAddIn
{
    [ComVisible(true)]
    public class AttachmentContextMenuRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ui;

        public string GetCustomUI(string ribbonID)
        {
            try { DevDiag.Log("GetCustomUI: ribbonID=" + (ribbonID ?? "<null>")); } catch { }
            DevDiag.Log("RibbonXML: forcing inline fallback XML (for custom image)");
            return GetInlineXmlFallback(); // ← בשלב הזה לא טוענים משאב מוטמע
        }



        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ui = ribbonUI;
            try { DevDiag.Log("RibbonXML: OnRibbonLoad hit"); } catch { }
        }
        // --- Inline/Signature helpers (paste inside AttachmentContextMenuRibbon class) ---
        private bool IsInlineImage(Outlook.Attachment att)
        {
            try
            {
                if (att == null) return false;

                string name = att.FileName ?? "";
                string ext = System.IO.Path.GetExtension(name).ToLowerInvariant();
                string cid = GetContentId(att); // null אם אין

                // HTML חלקי גוף
                if (ext == ".htm" || ext == ".html") return true;

                // תמונות—נטפל בהן כ-inline רק אם יש ContentId (כלומר משובצות בגוף HTML)
                if (ext == ".gif" || ext == ".jpg" || ext == ".jpeg" || ext == ".png")
                    return !string.IsNullOrEmpty(cid);

                // שאר הסוגים—נשמור כברירת מחדל
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


        // ===== כפתור בטאב Home (TabMail) =====
        public void OnHomeSaveButton(Office.IRibbonControl control)
        {
            try
            {
                DevDiag.Log("HomeBtn: click");

                Outlook.Application app = Globals.ThisAddIn.Application;

                // מייל פעיל/נבחר
                Outlook.MailItem mailItem = null;
                var insp = app.ActiveInspector();
                if (insp != null) mailItem = insp.CurrentItem as Outlook.MailItem;

                if (mailItem == null)
                {
                    var exp = app.ActiveExplorer();
                    var selection = exp != null ? exp.Selection : null;
                    if (selection != null && selection.Count > 0 && selection[1] is Outlook.MailItem mi)
                        mailItem = mi;
                }

                DevDiag.Log("HomeBtn: mailItem? " + (mailItem != null));
                if (mailItem == null) { MessageBox.Show("לא נבחר מייל."); return; }

                using (var dialog = new SaveEmailDialog())
                {
                    var dr = dialog.ShowDialog();
                    DevDiag.Log("HomeBtn: dialog result = " + dr);
                    if (dr != DialogResult.OK) return;

                    var ent = dialog.SelectedEntityInfo;
                    string entityName = dialog.SelectedEntityName ?? string.Empty;

                    string baseFolder = ResolveArchiveRoot();
                    string category = dialog.SelectedCategory ?? string.Empty;
                    string requestNumber = dialog.RequestNumber ?? string.Empty;
                    string categoryFolder = EnsureCategoryFolder(baseFolder, category);

                    DevDiag.Log("HomeBtn: baseFolder=" + baseFolder
                        + " | category='" + category + "'"
                        + " | categoryFolder=" + categoryFolder
                        + " | req='" + requestNumber + "'");

                    string ensureErr;
                    if (!EnsureWritableFolder(categoryFolder, out ensureErr))
                    {
                        DevDiag.Log("HomeBtn: EnsureWritableFolder FAIL: " + ensureErr);
                        MessageBox.Show("אין אפשרות לכתוב לתיקייה: " + categoryFolder + "\r\n" + ensureErr,
                            "שגיאת גישה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var writer = new ArchiveWriter();
                    DevDiag.Log("HomeBtn: ArchiveWriter created");
                    DevDiag.Log("HomeBtn: SelectedOption=" + dialog.SelectedOption);

                    if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveEmail)
                    {
                        // שם בסיס: מותאם אישית אם סומן, אחרת נושא המייל
                        string fileBase;
                        if (dialog.UseCustomFileName)
                        {
                            fileBase = CleanFileName(System.IO.Path.GetFileNameWithoutExtension(dialog.CustomFileName));
                            if (string.IsNullOrWhiteSpace(fileBase)) fileBase = "NoSubject";
                            DevDiag.Log($"HomeBtn: custom filename requested -> '{fileBase}'");
                        }
                        else
                        {
                            string subject = mailItem.Subject ?? "NoSubject";
                            fileBase = CleanFileName(subject);
                            if (string.IsNullOrWhiteSpace(fileBase)) fileBase = "NoSubject";
                            DevDiag.Log($"HomeBtn: using mail subject -> subject='{subject}', safe='{fileBase}'");
                        }

                        string initialPath = System.IO.Path.Combine(categoryFolder, fileBase + ".msg");
                        string fullPath = GetUniquePath(initialPath);

                        DevDiag.Log($"HomeBtn: saving MSG. path='{fullPath}' (initial='{initialPath}')");

                        // עדיף עם דיאגנוסטיקה
                        if (!TrySaveMsgWithDiag(mailItem, fullPath, out string saveErr))
                        {
                            DevDiag.Log("HomeBtn: SaveAs FAILED: " + saveErr);
                            MessageBox.Show("שמירת הקובץ נכשלה:\r\n" + saveErr, "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        DevDiag.Log("HomeBtn: saved msg exists? " + System.IO.File.Exists(fullPath));

                        try
                        {
                            bool ok = (ent != null)
                                ? writer.TryInsertRecord(ent, requestNumber, fullPath, System.IO.Path.GetFileName(fullPath))
                                : writer.TryInsertRecord(entityName, requestNumber, fullPath, System.IO.Path.GetFileName(fullPath));
                            DevDiag.Log("HomeBtn: DB insert after MSG ok? " + ok);
                        }
                        catch (Exception exDb)
                        {
                            DevDiag.Log("HomeBtn: DB insert after MSG EX: " + exDb.Message);
                        }

                        MessageBox.Show("המייל נשמר:\n" + fullPath, "שמירה",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    else if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveAttachmentsOnly
      || dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveAttachments)
                    {
                        if (mailItem.Attachments == null || mailItem.Attachments.Count == 0)
                        {
                            MessageBox.Show("אין קבצים מצורפים במייל.");
                            return;
                        }

                        // סינון inline
                        var validAttachments = new List<Outlook.Attachment>();
                        foreach (Outlook.Attachment att in mailItem.Attachments)
                            if (!IsInlineImage(att)) validAttachments.Add(att);

                        DevDiag.Log("HomeBtn: attachments total=" + mailItem.Attachments.Count + ", valid=" + validAttachments.Count);

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
                            int index = 0;
                            int saved = 0;

                            foreach (var att in validAttachments)
                            {
                                if (!selectedSet.Contains(att.FileName))
                                    continue;

                                index++;

                                string rawName = att.FileName ?? "";
                                string ext = System.IO.Path.GetExtension(rawName);
                                if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";

                                // בסיס השם: שם מותאם (עם מיספור אם יש יותר מקובץ אחד) או לפי שם המצורף
                                string baseForThis;
                                if (dialog.UseCustomFileName)
                                {
                                    string baseClean = CleanFileName(dialog.CustomFileName);
                                    baseForThis = (selCount == 1) ? baseClean : (baseClean + "_" + index);
                                }
                                else
                                {
                                    string baseName = System.IO.Path.GetFileNameWithoutExtension(rawName) ?? "attachment";
                                    baseForThis = CleanFileName(baseName);
                                    if (string.IsNullOrWhiteSpace(baseForThis)) baseForThis = "attachment";
                                }

                                string initialPath = System.IO.Path.Combine(categoryFolder, baseForThis + ext);
                                string filePath = GetUniquePath(initialPath);

                                DevDiag.Log($"HomeBtn: saving ATT. raw='{rawName}', base='{baseForThis}', ext='{ext}', path='{filePath}', UseCustom={dialog.UseCustomFileName}, selCount={selCount}, index={index}");

                                try
                                {
                                    att.SaveAsFile(filePath);
                                    DevDiag.Log("HomeBtn: saved att exists? " + System.IO.File.Exists(filePath));
                                    saved++;

                                    try
                                    {
                                        bool ok = (ent != null)
                                            ? writer.TryInsertRecord(ent, requestNumber, filePath, System.IO.Path.GetFileName(filePath))
                                            : writer.TryInsertRecord(entityName, requestNumber, filePath, System.IO.Path.GetFileName(filePath));
                                        DevDiag.Log("HomeBtn: DB insert after ATT ok? " + ok);
                                    }
                                    catch (Exception exDbA)
                                    {
                                        DevDiag.Log("HomeBtn: DB insert after ATT EX: " + exDbA.Message);
                                    }
                                }
                                catch (COMException comEx)
                                {
                                    DevDiag.Log("HomeBtn: att.SaveAsFile COMEX 0x" + comEx.HResult.ToString("X") + ": " + comEx.Message);
                                }
                                catch (Exception exAtt)
                                {
                                    DevDiag.Log("HomeBtn: att.SaveAsFile EX: " + exAtt.Message);
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
                try { DevDiag.Log("HomeBtn: ROOT COMEX 0x" + comRoot.HResult.ToString("X") + ": " + comRoot.Message); } catch { }
                MessageBox.Show("שגיאת COM:\r\n" + comRoot.Message);
            }
            catch (Exception ex)
            {
                try { DevDiag.Log("HomeBtn: EX " + ex.Message); } catch { }
                MessageBox.Show("שגיאה: " + ex.Message);
            }
        }

        // ===== קליק-ימני על מצורפים =====
        public bool GetVisibleForAttachment(Office.IRibbonControl control)
        {
            try
            {
                var sel = control != null ? control.Context as Outlook.AttachmentSelection : null;
                return sel != null && sel.Count > 0;
            }
            catch { return false; }
        }

        public void OnSaveAttachmentToArchive(Office.IRibbonControl control)
        {
            Outlook.AttachmentSelection sel = null;
            Outlook.MailItem mail = null;

            var selectedAttachments = new List<Outlook.Attachment>();
            var selectedIndices = new List<int>();

            try
            {
                sel = control != null ? control.Context as Outlook.AttachmentSelection : null;
                if (sel == null || sel.Count == 0) return;

                mail = sel.Parent as Outlook.MailItem;

                // אוסף את העצמים שנבחרו (כדי לדעת מי ה-preselected)
                var enumerable = sel as IEnumerable;
                if (enumerable != null)
                {
                    foreach (object o in enumerable)
                    {
                        var att = o as Outlook.Attachment;
                        if (att != null) selectedAttachments.Add(att);
                    }
                }

                // ממפה אינדקסים פנימיים של המצורפים הנבחרים לתוך mail.Attachments
                if (mail != null && mail.Attachments != null && mail.Attachments.Count > 0)
                {
                    foreach (var selAtt in selectedAttachments)
                    {
                        int idx = FindAttachmentIndexInMail(mail, selAtt);
                        if (idx > 0) selectedIndices.Add(idx);
                    }
                }

                // פותח את הדיאלוג עם preselect (כפי שעשית)
                using (var dlg = new SaveEmailDialog(mail, selectedIndices))
                {
                    var dr = dlg.ShowDialog();
                    if (dr != DialogResult.OK) return;

                    var ent = dlg.SelectedEntityInfo;
                    string entityName = dlg.SelectedEntityName ?? string.Empty;
                    string category = dlg.SelectedCategory ?? string.Empty;
                    string requestNumber = dlg.RequestNumber ?? string.Empty;

                    string baseFolder = ResolveArchiveRoot();
                    string categoryFolder = EnsureCategoryFolder(baseFolder, category);

                    if (!EnsureWritableFolder(categoryFolder, out string ensureErr))
                    {
                        DevDiag.Log("CtxBtn: EnsureWritableFolder FAIL: " + ensureErr);
                        MessageBox.Show("אין אפשרות לכתוב לתיקייה: " + categoryFolder + "\r\n" + ensureErr,
                            "שגיאת גישה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var writer = new ArchiveWriter();
                    DevDiag.Log($"CtxBtn: dialog ok. Option={dlg.SelectedOption}, UseCustom={dlg.UseCustomFileName}, Custom='{dlg.CustomFileName}'");

                    if (dlg.SelectedOption == SaveEmailDialog.SaveOption.SaveEmail)
                    {
                        // שם קובץ: מותאם אם סומן, אחרת נושא המייל
                        string fileBase = dlg.UseCustomFileName
                            ? CleanFileName(dlg.CustomFileName)
                            : CleanFileName(mail.Subject ?? "NoSubject");
                        if (string.IsNullOrWhiteSpace(fileBase)) fileBase = "NoSubject";

                        string fullPath = GetUniquePath(System.IO.Path.Combine(categoryFolder, fileBase + ".msg"));
                        DevDiag.Log($"CtxBtn: saving MSG. base='{fileBase}', path='{fullPath}'");

                        string err;
                        if (!TrySaveMsgWithDiag(mail, fullPath, out err))
                        {
                            DevDiag.Log("CtxBtn: SaveAs MSG FAIL: " + err);
                            MessageBox.Show("נכשלה שמירת המייל:\r\n" + err, "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        DevDiag.Log("CtxBtn: saved msg exists? " + System.IO.File.Exists(fullPath));

                        try
                        {
                            bool ok = (ent != null)
                                ? writer.TryInsertRecord(ent, requestNumber, fullPath, System.IO.Path.GetFileName(fullPath))
                                : writer.TryInsertRecord(entityName, requestNumber, fullPath, System.IO.Path.GetFileName(fullPath));
                            DevDiag.Log("CtxBtn: DB insert after MSG ok? " + ok);
                        }
                        catch (Exception exDb) { DevDiag.Log("CtxBtn: DB insert after MSG EX: " + exDb.Message); }

                        MessageBox.Show("המייל נשמר:\n" + fullPath, "שמירה", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else // SaveAttachments / SaveAttachmentsOnly
                    {
                        // המצורפים שנבחרו בקליק-ימני כבר נמצאים ב-selectedAttachments
                        if (selectedAttachments.Count == 0)
                        {
                            MessageBox.Show("לא נבחרו קבצים מצורפים.", "מידע", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        int selCount = selectedAttachments.Count;
                        int saved = 0;

                        for (int i = 0; i < selCount; i++)
                        {
                            var att = selectedAttachments[i];

                            string rawName = att.FileName ?? "";
                            string ext = System.IO.Path.GetExtension(rawName);
                            if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";

                            string baseForThis;
                            if (dlg.UseCustomFileName)
                            {
                                string baseClean = CleanFileName(dlg.CustomFileName);
                                baseForThis = (selCount == 1) ? baseClean : (baseClean + "_" + (i + 1));
                            }
                            else
                            {
                                string baseName = System.IO.Path.GetFileNameWithoutExtension(rawName) ?? "attachment";
                                baseForThis = CleanFileName(baseName);
                                if (string.IsNullOrWhiteSpace(baseForThis)) baseForThis = "attachment";
                            }

                            string filePath = GetUniquePath(System.IO.Path.Combine(categoryFolder, baseForThis + ext));
                            DevDiag.Log($"CtxBtn: saving ATT. raw='{rawName}', base='{baseForThis}', ext='{ext}', path='{filePath}', UseCustom={dlg.UseCustomFileName}, idx={i + 1}/{selCount}");

                            try
                            {
                                att.SaveAsFile(filePath);
                                saved++;

                                try
                                {
                                    bool ok = (ent != null)
                                        ? writer.TryInsertRecord(ent, requestNumber, filePath, System.IO.Path.GetFileName(filePath))
                                        : writer.TryInsertRecord(entityName, requestNumber, filePath, System.IO.Path.GetFileName(filePath));
                                    DevDiag.Log("CtxBtn: DB insert after ATT ok? " + ok);
                                }
                                catch (Exception exDbA) { DevDiag.Log("CtxBtn: DB insert after ATT EX: " + exDbA.Message); }
                            }
                            catch (System.Runtime.InteropServices.COMException comEx)
                            {
                                DevDiag.Log("CtxBtn: att.SaveAsFile COMEX 0x" + comEx.HResult.ToString("X") + ": " + comEx.Message);
                            }
                            catch (Exception exAtt)
                            {
                                DevDiag.Log("CtxBtn: att.SaveAsFile EX: " + exAtt.Message);
                            }
                        }

                        MessageBox.Show(
                            saved > 0 ? "הקבצים שנבחרו נשמרו בהצלחה." : "שמירה נכשלה.",
                            "שמירת קבצים", MessageBoxButtons.OK,
                            saved > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                    }
                }

            }
            finally
            {
                if (selectedAttachments != null)
                {
                    foreach (var a in selectedAttachments)
                        if (a != null) { try { Marshal.ReleaseComObject(a); } catch { } }
                }
                if (mail != null) { try { Marshal.ReleaseComObject(mail); } catch { } }
                if (sel != null) { try { Marshal.ReleaseComObject(sel); } catch { } }
            }
        }


        // ===== עזרי אינדקס מצורף =====
        private static int FindAttachmentIndexInMail(Outlook.MailItem mail, Outlook.Attachment selected)
        {
            if (mail == null || mail.Attachments == null || selected == null) return 0;

            string selName = SafeGetFileName(selected);
            int selSize = SafeGetSize(selected);

            Outlook.Attachments atts = null;
            try
            {
                atts = mail.Attachments;
                int count = atts != null ? atts.Count : 0;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Attachment att = null;
                    try
                    {
                        att = atts[i];
                        string name = SafeGetFileName(att);
                        int size = SafeGetSize(att);

                        if (!string.IsNullOrEmpty(selName) &&
                            selName.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            if (selSize > 0 && size > 0)
                            {
                                if (selSize == size) return i;
                            }
                            else
                            {
                                return i;
                            }
                        }
                    }
                    finally
                    {
                        if (att != null) { try { Marshal.ReleaseComObject(att); } catch { } }
                    }
                }
            }
            finally
            {
                if (atts != null) { try { Marshal.ReleaseComObject(atts); } catch { } }
            }
            return 0;
        }

        private static string SafeGetFileName(Outlook.Attachment att)
        {
            try { return att != null ? att.FileName : null; } catch { return null; }
        }
        private static int SafeGetSize(Outlook.Attachment att)
        {
            try { return att != null ? att.Size : 0; } catch { return 0; }
        }

        // ===== עזרי שמירה/תיקיות =====
        private bool EnsureWritableFolder(string folder, out string error)
        {
            error = null;
            try
            {
                Directory.CreateDirectory(folder);
                string probe = System.IO.Path.Combine(folder, "~write_probe_" + Guid.NewGuid().ToString("N") + ".tmp");
                using (var fs = File.Create(probe)) { }
                File.Delete(probe);
                return true;
            }
            catch (Exception ex) { error = ex.Message; return false; }
        }

        private string GetUniquePath(string initialPath)
        {
            if (string.IsNullOrWhiteSpace(initialPath))
                throw new ArgumentException("initialPath is null/empty");

            string dir = System.IO.Path.GetDirectoryName(initialPath);
            string name = System.IO.Path.GetFileNameWithoutExtension(initialPath);
            string ext = System.IO.Path.GetExtension(initialPath);

            if (string.IsNullOrWhiteSpace(dir))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (ext == null) ext = string.Empty;

            try { Directory.CreateDirectory(dir); } catch { }

            string candidate = System.IO.Path.Combine(dir, name + ext);
            if (!File.Exists(candidate)) return candidate;

            for (int i = 1; i < 1000; i++)
            {
                candidate = System.IO.Path.Combine(dir, name + "_" + i + ext);
                if (!File.Exists(candidate)) return candidate;
            }

            candidate = System.IO.Path.Combine(dir, name + "_" +
                       DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ext);
            return candidate;
        }

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
                root = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SavedMails");

            if (!System.IO.Path.IsPathRooted(root))
            {
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                root = System.IO.Path.GetFullPath(System.IO.Path.Combine(docs, root));
            }

            try { Directory.CreateDirectory(root); } catch { }

            try { DevDiag.Log("HomeBtn: ResolveArchiveRoot -> " + root); } catch { }
            return root;
        }

        private string EnsureCategoryFolder(string baseFolder, string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                Directory.CreateDirectory(baseFolder);
                return baseFolder;
            }
            string full = System.IO.Path.Combine(baseFolder, CleanFileName(category));
            Directory.CreateDirectory(full);
            return full;
        }

        private string CleanFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Unnamed";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }

        // ===== XML helpers =====
        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new InvalidOperationException("Embedded resource not found: " + resourceName);
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public IPictureDisp LoadImage(string imageId)
        {
            try
            {
                // ניסיון דרך Resources לפי שם (כולל BarAddin_icon)
                object obj = global::BarOutlookAddIn.Properties.Resources.ResourceManager.GetObject(imageId);

                if (obj is Icon ico) return PictureDispConverter.ToIPictureDisp(ico.ToBitmap());
                if (obj is Bitmap bmp) return PictureDispConverter.ToIPictureDisp(bmp);
                if (obj is Image img) return PictureDispConverter.ToIPictureDisp(img);

                // אם בטעות הוגדר כמחרוזת/נתיב – נטען מהדיסק
                if (obj is string pathStr && File.Exists(pathStr))
                {
                    using (var im = Image.FromFile(pathStr))
                        return PictureDispConverter.ToIPictureDisp(new Bitmap(im));
                }

                return null;
            }
            catch { return null; }
        }

        // ממיר Image -> IPictureDisp עבור RibbonX
        private class PictureDispConverter : System.Windows.Forms.AxHost
        {
            private PictureDispConverter() : base("") { }
            public static IPictureDisp ToIPictureDisp(Image image)
                => (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        // בתוך המחלקה AttachmentContextMenuRibbon
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
                try { DevDiag.Log("SaveAs COMEX: " + error); } catch { }
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { DevDiag.Log("SaveAs EX: " + error); } catch { }
                return false;
            }
        }


        private static string GetInlineXmlFallback()
        {
            return
        @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui""
          onLoad=""OnRibbonLoad""
          loadImage=""LoadImage"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""grpBarArchive"" label=""שמירה לבר"">
          <button id=""btnSaveSelectedEmail""
                  label=""שמירה לארכיב""
                  size=""large""
                  image=""BarAddin_icon""
                  onAction=""OnHomeSaveButton"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuAttachments"">
      <button id=""btnSaveToArchive""
              label=""שמירה לארכיב""
              image=""BarAddin_icon""
              onAction=""OnSaveAttachmentToArchive""
              getVisible=""GetVisibleForAttachment""/>
    </contextMenu>
  </contextMenus>
</customUI>";
        }


    }
}
