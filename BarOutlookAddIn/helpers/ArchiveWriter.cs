using System;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using BarOutlookAddIn;


namespace BarOutlookAddIn.Helpers
{
    public class ArchiveWriter
    {
        // -------- Preferred overload: uses EntityInfo so we pass correct system type + definement --------
        public bool TryInsertRecord(EntityInfo entity, string dspEntityNum, string fullPath, string fileDesc)
        {
            if (entity == null) return false;

            string cs = GetConnectionString();
            if (string.IsNullOrWhiteSpace(cs)) return false;

            try
            {
                // Ensure we have a proper system type char (e.g., 'ת', 'ב', 'פ', ...)
                char entityTypeChar = 'ב';
                if (!string.IsNullOrWhiteSpace(entity.SystemType))
                {
                    string st = entity.SystemType.Trim();
                    if (st.Length > 0) entityTypeChar = st[0];
                }

                // IMPORTANT: definement must come from EntityInfo (e.g., 10 for תוכנית)
                int definement = entity.Definement;

                // For תוכנית – DO NOT clean; keep full text including slashes/hyphens.
                // For others – keep legacy cleaning behavior.
                string cleanedDsp = CleanDspForEntity(entityTypeChar, dspEntityNum);

                using (var con = new SqlConnection(cs))
                using (var cmd = new SqlCommand("SP_Insert_Archive", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                    cmd.Parameters.Add("@entity_type", SqlDbType.NVarChar, 1).Value = entityTypeChar.ToString(); // NVARCHAR for Hebrew
                    cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = definement;

                    // Use NVARCHAR for Hebrew/UTF-8 strings
                    cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.NVarChar, 100).Value = (object)(cleanedDsp ?? "") ?? "";
                    cmd.Parameters.Add("@File_Name", SqlDbType.NVarChar, 255).Value = (object)(fileDesc ?? "") ?? "";
                    cmd.Parameters.Add("@File_Location", SqlDbType.NVarChar, -1).Value = (object)(fullPath ?? "") ?? ""; // -1 = NVARCHAR(MAX)

                    con.Open();
                    int rows = cmd.ExecuteNonQuery();

                    // Log runnable EXEC for diagnostics
                    try { DevDiag.Log("EXEC " + DevDiag.AsExec(cmd) + " | rows=" + rows); } catch { }
                    
                    return rows > 0;
                }
            }
            catch (Exception ex)
            {
                try { DevDiag.Log("DB EX: " + ex.Message); } catch { }
                return false;
            }
        }

        // -------- Legacy-compatible overload (string entityName) --------
        public bool TryInsertRecord(string entityName, string dspEntityNum, string fullPath, string fileDesc)
        {
            string cs = GetConnectionString();
            if (string.IsNullOrWhiteSpace(cs)) return false;

            try
            {
                string systemEntityType = GetSystemEntityType(cs, entityName); // e.g. "ת"/"ב"/"פ"
                char entityTypeChar = (!string.IsNullOrWhiteSpace(systemEntityType)) ? systemEntityType.Trim()[0] : 'ב';

                int definement;
                if (entityTypeChar == 'ת')
                {
                    // For תוכנית: try to fetch definement from System_Entity by Description_Code
                    definement = TryGetPlanDefinementFromSystemEntity(cs, entityName);
                }
                else
                {
                    // For others: from Entity_Type_Control
                    definement = GetDefinementEntityType(cs, entityName);
                }

                string cleanedDsp = (entityTypeChar == 'ת')
                    ? (dspEntityNum ?? "")               // keep as-is for תוכנית
                    : CleanDspIfNeeded(dspEntityNum);    // legacy cleaning for others

                using (var con = new SqlConnection(cs))
                using (var cmd = new SqlCommand("SP_Insert_Archive", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                    cmd.Parameters.Add("@entity_type", SqlDbType.NVarChar, 1).Value = entityTypeChar.ToString();
                    cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = definement;
                    cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.NVarChar, 100).Value = (object)(cleanedDsp ?? "") ?? "";
                    cmd.Parameters.Add("@File_Name", SqlDbType.NVarChar, 255).Value = (object)(fileDesc ?? "") ?? "";
                    cmd.Parameters.Add("@File_Location", SqlDbType.NVarChar, -1).Value = (object)(fullPath ?? "") ?? "";

                    con.Open();
                    int rows = cmd.ExecuteNonQuery();

                    try { DevDiag.Log("EXEC " + DevDiag.AsExec(cmd) + " | rows=" + rows); } catch { }

                    return rows > 0;
                }
            }
            catch (Exception ex)
            {
                try { DevDiag.Log("DB EX(legacy): " + ex.Message); } catch { }
                return false;
            }
        }

        // ---------------- helpers ----------------

        private string GetConnectionString()
        {
            try
            {
                var props = Properties.Settings.Default.Properties;
                if (props != null && props["ConnectionString"] != null)
                {
                    var v = Properties.Settings.Default["ConnectionString"] as string;
                    if (!string.IsNullOrWhiteSpace(v)) return v;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// For תוכנית ('ת'): return dsp as-is (supports "6870/6-מזרחי").
        /// For others: legacy cleaning (digits extraction unless IsShomron).
        /// </summary>
        private string CleanDspForEntity(char entityTypeChar, string dsp)
        {
            if (entityTypeChar == 'ת')
                return dsp ?? "";

            return CleanDspIfNeeded(dsp);
        }

        // Legacy behavior used for non-plan entities
        private string CleanDspIfNeeded(string dsp)
        {
            try
            {
                bool isShomron = false;
                try
                {
                    var props = Properties.Settings.Default.Properties;
                    if (props != null && props["IsShomron"] != null)
                    {
                        var v = Properties.Settings.Default["IsShomron"];
                        if (v is bool) isShomron = (bool)v;
                    }
                }
                catch { }

                if (isShomron) return dsp ?? "";
                if (string.IsNullOrWhiteSpace(dsp)) return "";

                string[] parts = Regex.Split(dsp, @"\D+");
                for (int i = 0; i < parts.Length; i++)
                {
                    string p = parts[i];
                    if (!string.IsNullOrEmpty(p) && p.Length > 1) return p;
                }
                return dsp;
            }
            catch { return dsp ?? ""; }
        }

        private string GetSystemEntityType(string cs, string entityName)
        {
            if (string.IsNullOrWhiteSpace(entityName)) return "ב";
            string sysType = null;

            using (var con = new SqlConnection(cs))
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT TOP 1 [system_entity_type] " +
                    "FROM [dbo].[Entity_Type_Control] " +
                    "WHERE [entity_name]=@name";
                cmd.Parameters.AddWithValue("@name", entityName); // NVARCHAR
                con.Open();
                object o = cmd.ExecuteScalar();
                if (o != null && o != DBNull.Value) sysType = o.ToString();
            }

            return string.IsNullOrWhiteSpace(sysType) ? "ב" : sysType;
        }

        private int GetDefinementEntityType(string cs, string entityName)
        {
            if (string.IsNullOrWhiteSpace(entityName)) return 0;
            int code = 0;

            using (var con = new SqlConnection(cs))
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT TOP 1 [entity_type] " +
                    "FROM [dbo].[Entity_Type_Control] " +
                    "WHERE [entity_name]=@name";
                cmd.Parameters.AddWithValue("@name", entityName); // NVARCHAR
                con.Open();
                object o = cmd.ExecuteScalar();
                if (o != null && o != DBNull.Value)
                {
                    int.TryParse(o.ToString(), out code);
                }
            }
            return code;
        }

        /// <summary>
        /// For תוכנית: try to get definement from System_Entity table (definement_entity_type)
        /// by Description_Code (entityName). If not found, return 0.
        /// </summary>
        private int TryGetPlanDefinementFromSystemEntity(string cs, string entityName)
        {
            if (string.IsNullOrWhiteSpace(entityName)) return 0;

            int def = 0;
            using (var con = new SqlConnection(cs))
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT TOP 1 [definement_entity_type] " +
                    "FROM [dbo].[System_Entity] " +
                    "WHERE [Code_Identification] = N'ת' AND [Description_Code] = @name";
                cmd.Parameters.AddWithValue("@name", entityName); // NVARCHAR
                con.Open();
                object o = cmd.ExecuteScalar();
                if (o != null && o != DBNull.Value)
                {
                    int.TryParse(o.ToString(), out def);
                }
            }
            return def;
        }
    }
}
