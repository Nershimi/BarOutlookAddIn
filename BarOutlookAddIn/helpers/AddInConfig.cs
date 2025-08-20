using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace BarOutlookAddIn.Helpers
{
    public class AddInConfig
    {
        // שמרתי תאימות אם בעתיד יהיה <Categories> ישן
        public List<string> Categories { get; } = new List<string>();
        public string DefaultCategory { get; private set; } = "";

        // חדשים לפי הסכֵמה ששלחת
        public string ArchivePath { get; private set; } = "";       // NB_Archive_Path
        public string DefaultSystem { get; private set; } = "";      // DefaultSystem
        public string DefaultEntity { get; private set; } = "";      // DefaultEntity
        public bool Numerator { get; private set; } = false;         // numerator
        public bool Shomron { get; private set; } = false;           // Shomron

        // פרטי SQL (אם תרצה להשתמש)
        public string SqlServerAddress { get; private set; } = "";   // NB_Server_Address
        public string SqlDbName { get; private set; } = "";          // NB_DB_Name
        public string SqlUserName { get; private set; } = "";        // SqlUserName
        public string SqlPassword { get; private set; } = "";        // SqlPassword

        public static AddInConfig Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path is empty.");

            if (!File.Exists(path))
                throw new FileNotFoundException("Settings file not found.", path);

            XDocument xdoc = XDocument.Load(path);
            AddInConfig cfg = new AddInConfig();

            XElement root = xdoc.Root;
            if (root == null) throw new InvalidDataException("Missing root element <Settings>.");

            // --- תמיכה אחורה: <Categories> (לא חובה בסכֵמה החדשה)
            XElement catsElem = root.Element("Categories");
            if (catsElem != null)
            {
                List<string> cats = catsElem.Elements("Category")
                    .Select(e => e.Attribute("name") != null ? e.Attribute("name").Value : "")
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct()
                    .ToList();

                if (cats.Count > 0)
                {
                    cfg.Categories.AddRange(cats);
                }

                XElement defCat = root.Element("DefaultCategory");
                if (defCat != null) cfg.DefaultCategory = (defCat.Value ?? "").Trim();
            }

            // --- סכֵמת NB החדשה
            XElement nb = root.Element("NB");
            if (nb != null)
            {
                cfg.SqlServerAddress = GetValue(nb, "NB_Server_Address");
                cfg.ArchivePath = GetValue(nb, "NB_Archive_Path");
                cfg.SqlDbName = GetValue(nb, "NB_DB_Name");
                cfg.SqlUserName = GetValue(nb, "SqlUserName");
                cfg.SqlPassword = GetValue(nb, "SqlPassword");
            }

            cfg.DefaultSystem = GetValue(root, "DefaultSystem");
            cfg.DefaultEntity = GetValue(root, "DefaultEntity");
            cfg.Numerator = ParseBool(GetValue(root, "numerator"));
            cfg.Shomron = ParseBool(GetValue(root, "Shomron"));

            return cfg;
        }

        private static string GetValue(XElement parent, string name)
        {
            XElement el = parent.Element(name);
            return el != null ? (el.Value ?? "").Trim() : "";
        }

        private static bool ParseBool(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            bool b;
            if (bool.TryParse(s.Trim(), out b)) return b;

            // תמיכה בערכים לא תקניים (0/1, yes/no)
            string t = s.Trim().ToLowerInvariant();
            if (t == "1" || t == "yes" || t == "y" || t == "true") return true;
            return false;
        }
    }
}
