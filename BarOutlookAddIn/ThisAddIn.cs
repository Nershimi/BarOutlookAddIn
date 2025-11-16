using System;
using System.IO;
using System.Xml.Linq;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;

namespace BarOutlookAddIn
{
    public partial class ThisAddIn
    {
        // מציגים טוסט חיבור למסד פעם אחת בלבד לכל סשן
        private static bool _dbPingShown = false;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 1) Init log folder FIRST
            DevDiag.ConfigureLogFolder(@"C:\bar\m9");
            DevDiag.Log("Startup: entered. Using log at: " + DevDiag.CurrentLogPath);

            // 2) Only then start the rest, so early logs ייכנסו לנתיב שבחרת
            ShowDbConnectionToastIfNeeded();
        }
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try { DevDiag.Log("CreateRibbonExtensibilityObject: returning AttachmentContextMenuRibbon"); } catch { }
            return new AttachmentContextMenuRibbon();
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

        private void LogSpInsertArchiveSignature()
        {
            try
            {
                var settings = global::BarOutlookAddIn.Properties.Settings.Default;
                var cs = settings.Properties != null && settings["ConnectionString"] != null
                    ? settings["ConnectionString"] as string
                    : null;

                DevDiag.Log("DbInspector: checking SP_Insert_Archive. cs exists? " + (!string.IsNullOrWhiteSpace(cs)));

                if (string.IsNullOrWhiteSpace(cs))
                {
                    DevDiag.Log("DbInspector: no __ConnectionString__ available.");
                    return;
                }

                string proc = "dbo.SP_Insert_Archive";

                // 1) existence
                bool exists = Helpers.DbInspector.StoredProcedureExists(cs, proc);
                DevDiag.Log("DbInspector: StoredProcedureExists(" + proc + ") = " + exists);

                // 2) list actual parameters
                var actual = Helpers.DbInspector.GetStoredProcedureParameters(cs, proc);
                DevDiag.Log("DbInspector: actual parameters count = " + actual.Count);
                for (int i = 0; i < actual.Count; i++)
                    DevDiag.Log("DbInspector: param[" + (i + 1) + "] = " + actual[i]);

                // 3) quick compare against expected signature
                var expected = new List<string>
                {
                    "@Estate_Number bigint",
                    "@entity_type char(1)",
                    "@definement_entity_type int",
                    "@Org_Entity_Number varchar(50)",
                    "@File_Name varchar(300)",
                    "@File_Location varchar(300)"
                };

                if (Helpers.DbInspector.MatchesExpectedSignature(cs, proc, expected, out string message))
                {
                    DevDiag.Log("DbInspector: Signature matches expected.");
                }
                else
                {
                    DevDiag.Log("DbInspector: Signature DOES NOT match. " + message);
                }
            }
            catch (Exception ex)
            {
                DevDiag.Log("DbInspector: check EX " + ex.Message);
            }
        }

        // ------- חשוב: InternalStartup חייב להיות כאן, באותו namespace, כדי ש-Designer יקרא לו -------
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
