using System.Collections.Generic;

namespace BarOutlookAddIn.Helpers
{
    // Maps one-letter system type codes to a Hebrew description.
    // Use EntityTypeCatalog.GetDescription(code) to convert a system type code ("ב", "פ", "ת", etc.)
    // to its human-readable description. If code is unknown or null/empty the code itself or an empty
    // string is returned (safe fallback).
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
