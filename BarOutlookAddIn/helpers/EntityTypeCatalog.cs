using System.Collections.Generic;

namespace BarOutlookAddIn.Helpers
{
    // Map one-letter system type code to a Hebrew description.
    internal static class EntityTypeCatalog
    {
        private static readonly Dictionary<string, string> _map =
            new Dictionary<string, string>
            {
                {"ב", "בקשות"},
                {"פ", "פיקוח"},
                {"ת", "ישות תכנונית"},
                {"כ", "ישות כללית"},
                {"ע", "תביעה"}
            };

        public static string GetDescription(string code)
        {
            if (string.IsNullOrEmpty(code)) return "";
            string desc;
            return _map.TryGetValue(code, out desc) ? desc : code;
        }
    }
}
