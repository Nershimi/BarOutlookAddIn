using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BarOutlookAddIn
{
    public partial class SaveEmailRibbon
    {
        private void SaveEmailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // ניתן להסיר את ההודעה אם לא נדרש
        }

        private void btnSaveSelectedEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.Selection selection = app.ActiveExplorer().Selection;

                if (selection.Count > 0 && selection[1] is Outlook.MailItem mailItem)
                {
                    SaveEmailDialog dialog = new SaveEmailDialog();
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveFolder = @"C:\SavedEmails";
                        Directory.CreateDirectory(saveFolder);

                        string subject = CleanFileName(mailItem.Subject ?? "NoSubject");
                        string timestamp = mailItem.ReceivedTime.ToString("yyyy-MM-dd_HH-mm-ss");

                        if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveEmail)
                        {
                            string fileName = $"{timestamp}_{subject}.msg";
                            string fullPath = Path.Combine(saveFolder, fileName);
                            mailItem.SaveAs(fullPath, Outlook.OlSaveAsType.olMSG);
                            MessageBox.Show("המייל נשמר:\n" + fullPath);
                        }
                        else if (dialog.SelectedOption == SaveEmailDialog.SaveOption.SaveAttachmentsOnly)
                        {
                            if (mailItem.Attachments.Count == 0)
                            {
                                MessageBox.Show("אין קבצים מצורפים במייל.");
                                return;
                            }

                            // סינון קבצים לא רצויים (תמונות חתימה וכדומה)
                            var validAttachments = new List<Outlook.Attachment>();
                            foreach (Outlook.Attachment att in mailItem.Attachments)
                            {
                                if (!IsInlineImage(att))
                                {
                                    validAttachments.Add(att);
                                }
                            }

                            if (validAttachments.Count == 0)
                            {
                                MessageBox.Show("לא נמצאו קבצים תקינים לשמירה.");
                                return;
                            }

                            // הצגת תיבת בחירה
                            var names = validAttachments.Select(a => a.FileName).ToList();
                            var selectionDialog = new AttachmentSelectionDialog(names);
                            if (selectionDialog.ShowDialog() == DialogResult.OK)
                            {
                                foreach (var att in validAttachments)
                                {
                                    if (selectionDialog.SelectedAttachments.Contains(att.FileName))
                                    {
                                        string filePath = Path.Combine(saveFolder, CleanFileName(att.FileName));
                                        att.SaveAsFile(filePath);
                                    }
                                }
                                MessageBox.Show("הקבצים שנבחרו נשמרו בהצלחה.");
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("לא נבחר מייל.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("שגיאה: " + ex.Message);
            }
        }

        private bool IsInlineImage(Outlook.Attachment attachment)
        {
            // קבצים מוטמעים (כמו חתימות) הם בדרך כלל inline עם CID
            try
            {
                return !string.IsNullOrEmpty(attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")?.ToString());
            }
            catch
            {
                return false;
            }
        }

        private string CleanFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
