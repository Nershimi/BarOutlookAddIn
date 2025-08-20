using System;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace BarOutlookAddIn
{
    public partial class ThisAddIn
    {
        // מציגים טוסט חיבור למסד פעם אחת בלבד לכל סשן
        private static bool _dbPingShown = false;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 1) Init log folder FIRST
            DevDiag.ConfigureLogFolder(@"C:\bar\logs");
            DevDiag.Log("Startup: entered. Using log at: " + DevDiag.CurrentLogPath);

            // 2) Only then start the rest, so early logs ייכנסו לנתיב שבחרת
            ShowDbConnectionToastIfNeeded();
        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Outlook כמעט ולא מרים אירוע זה בגרסאות חדשות; אפשר להשאיר ריק.
        }

       

        // ----------------- DB Toast Helpers -----------------

        private void ShowDbConnectionToastIfNeeded()
        {
            DevDiag.Log("DBToast: start");

            if (_dbPingShown) return;

            // 1) ננסה להביא ConnectionString מ-Settings של הפרויקט הזה
            string cs = GetConnectionStringFromSettings();

            // 2) אם אין, ננסה לבנות מתוך ה-XML ששמור ב-ConfigPath
            if (string.IsNullOrWhiteSpace(cs))
                cs = TryBuildConnectionStringFromXml();

            //string err;
            //if (TryPingDb(cs, out err))
            //{
            //    _dbPingShown = true;
            //    MessageBox.Show("החיבור למסד הנתונים הצליח.", "התחברות", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //else
            //{
            //    // אם תרצה להראות גם כשלון, בטל הערה:
            //    // MessageBox.Show("החיבור למסד הנתונים נכשל:\r\n" + (err ?? "לא נמצאה מחרוזת התחברות תקפה"),
            //    //                 "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private string GetConnectionStringFromSettings()
        {
            try
            {
                DevDiag.Log("GetCS: trying settings");
                // שימוש מפורש במרחב השמות כדי למנוע שגיאות 'Properties'
                var settings = global::BarOutlookAddIn.Properties.Settings.Default;
                var props = settings.Properties;
                if (props != null && props["ConnectionString"] != null)
                {
                    var val = settings["ConnectionString"];
                    return val as string;
                }
            }
            catch { }
            return null;
        }

        private string TryBuildConnectionStringFromXml()
        {
            try
            {
                var settings = global::BarOutlookAddIn.Properties.Settings.Default;

                // read ConfigPath first
                string cfgPath = null;
                var props = settings.Properties;
                if (props != null && props["ConfigPath"] != null)
                    cfgPath = settings["ConfigPath"] as string;

                DevDiag.Log("BuildCS: from XML start. ConfigPath=" + (cfgPath ?? "<null>") +
                            ", exists=" + (!string.IsNullOrWhiteSpace(cfgPath) && File.Exists(cfgPath)));

                if (string.IsNullOrWhiteSpace(cfgPath) || !File.Exists(cfgPath))
                    return null;

                XDocument xdoc = XDocument.Load(cfgPath);
                XElement root = xdoc.Root;
                XElement nb = root != null ? root.Element("NB") : null;
                if (nb == null) return null;

                string server = SafeElem(nb, "NB_Server_Address");
                string db = SafeElem(nb, "NB_DB_Name");
                string user = SafeElem(nb, "SqlUserName");
                string pass = SafeElem(nb, "SqlPassword");

                if (string.IsNullOrWhiteSpace(server) || string.IsNullOrWhiteSpace(db))
                    return null;

                string cs = "Server=" + server + ";Database=" + db + ";User Id=" + user + ";Password=" + pass + ";";
                cs = EnsureAppName(cs); // optional but recommended for tracing

                try
                {
                    if (props != null && props["ConnectionString"] != null)
                    {
                        settings["ConnectionString"] = cs;
                        settings.Save();
                    }
                }
                catch { }

                DevDiag.Log("BuildCS: final cs=" + cs);
                return cs;
            }
            catch (Exception ex)
            {
                DevDiag.Log("BuildCS: EX " + ex.Message);
                return null;
            }
        }


        private static string SafeElem(XElement parent, string name)
        {
            XElement el = parent.Element(name);
            return el != null ? (el.Value ?? "").Trim() : "";
        }
        private string EnsureAppName(string cs)
        {
            if (string.IsNullOrWhiteSpace(cs)) return cs;
            if (cs.IndexOf("Application Name=", StringComparison.OrdinalIgnoreCase) >= 0)
                return cs;
            if (!cs.EndsWith(";")) cs += ";";
            return cs + "Application Name=BarOutlookAddIn;";
        }

        // ------- חשוב: InternalStartup חייב להיות כאן, באותו namespace, כדי ש-Designer יקרא לו -------
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
