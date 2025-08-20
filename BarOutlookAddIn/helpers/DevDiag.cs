using System;
using System.IO;
using System.Text;
using System.Data.SqlClient;

namespace BarOutlookAddIn
{
    internal static class DevDiag
    {
        private static string _customRoot;

        public static void ConfigureLogFolder(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                _customRoot = null;
                return;
            }
            try
            {
                Directory.CreateDirectory(folder);
                _customRoot = folder;
            }
            catch
            {
                _customRoot = null; // fallback later
            }
        }

        private static string ResolveRoot()
        {
            // 1) runtime override
            if (!string.IsNullOrEmpty(_customRoot))
                return _customRoot;

            // 2) user setting (optional)
            try
            {
                var props = Properties.Settings.Default.Properties;
                if (props != null && props["LogFolder"] != null)
                {
                    var v = Properties.Settings.Default["LogFolder"] as string;
                    if (!string.IsNullOrWhiteSpace(v))
                    {
                        Directory.CreateDirectory(v);
                        return v;
                    }
                }
            }
            catch { /* ignore */ }

            // 3) fallback: %AppData%\BarOutlookAddIn
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BarOutlookAddIn");
        }

        public static string CurrentLogPath
        {
            get { return Path.Combine(ResolveRoot(), "debug.log"); }
        }

        public static void Log(string msg)
        {
            try
            {
                string path = CurrentLogPath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.AppendAllText(
                    path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + msg + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { /* swallow */ }
        }

        // Render SqlCommand → EXEC text (handy for copy/paste to SSMS)
        public static string AsExec(SqlCommand cmd)
        {
            var sb = new StringBuilder("EXEC " + cmd.CommandText + " ");
            for (int i = 0; i < cmd.Parameters.Count; i++)
            {
                var p = cmd.Parameters[i];
                string name = p.ParameterName.TrimStart('@');
                string val;
                if (p.Value == null || p.Value == DBNull.Value)
                {
                    val = "NULL";
                }
                else if (p.SqlDbType == System.Data.SqlDbType.Int ||
                         p.SqlDbType == System.Data.SqlDbType.BigInt ||
                         p.SqlDbType == System.Data.SqlDbType.SmallInt ||
                         p.SqlDbType == System.Data.SqlDbType.TinyInt)
                {
                    val = p.Value.ToString();
                }
                else
                {
                    val = "N'" + p.Value.ToString().Replace("'", "''") + "'";
                }
                sb.Append("@" + name + "=" + val);
                if (i < cmd.Parameters.Count - 1) sb.Append(", ");
            }
            return sb.ToString();
        }
    }
}
