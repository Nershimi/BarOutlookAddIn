using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

namespace BarOutlookAddIn
{
    public class SettingsLoader
    {
        private readonly string _xmlPath;
        private XmlDocument _doc;

        public SettingsLoader(string xmlPath)
        {
            _xmlPath = xmlPath;
            LoadXml();
        }

        private void LoadXml()
        {
            try
            {
                _doc = new XmlDocument();
                _doc.Load(_xmlPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בטעינת קובץ ההגדרות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetSqlConnectionString()
        {
            try
            {
                string server = _doc.SelectSingleNode("/Settings/NB/NB_Server_Address")?.InnerText ?? "";
                string db = _doc.SelectSingleNode("/Settings/NB/NB_DB_Name")?.InnerText ?? "";
                string user = _doc.SelectSingleNode("/Settings/NB/SqlUserName")?.InnerText ?? "";
                string pass = _doc.SelectSingleNode("/Settings/NB/SqlPassword")?.InnerText ?? "";

                if (string.IsNullOrWhiteSpace(server) || string.IsNullOrWhiteSpace(db))
                    throw new InvalidOperationException("הגדרות התחברות חסרות.");

                return $"Server={server};Database={db};User Id={user};Password={pass};";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בבניית מחרוזת ההתחברות: {ex.Message}", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public string GetDefaultEntity()
        {
            return _doc.SelectSingleNode("/Settings/DefaultEntity")?.InnerText ?? "";
        }

        public string GetDefaultSystem()
        {
            return _doc.SelectSingleNode("/Settings/DefaultSystem")?.InnerText ?? "";
        }

        public Dictionary<string, string> GetFlatSettings()
        {
            var settings = new Dictionary<string, string>();
            try
            {
                XmlNodeList nodes = _doc.SelectNodes("/Settings/*");
                foreach (XmlNode node in nodes)
                {
                    if (node.HasChildNodes && node.FirstChild is XmlText)
                    {
                        settings[node.Name] = node.InnerText;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בקריאת הגדרות כלליות: {ex.Message}");
            }

            return settings;
        }
    }
}
