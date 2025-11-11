using System;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace BarOutlookAddIn.Helpers
{
    public class ArchiveWriter
    {
        // ========= API =========

        public bool TryInsertRecord(EntityInfo entity, string dspEntityNum, string fullPath, string fileDesc)
        {
            if (entity == null) return false;

            var cs = GetConnectionString();
            if (string.IsNullOrWhiteSpace(cs)) return false;

            try
            {
                // קביעת סוג יישות (עברית) כפי שמגיע מה-DB
                char sysChar = 'ב';
                if (!string.IsNullOrWhiteSpace(entity.SystemType))
                {
                    var st = entity.SystemType.Trim();
                    if (st.Length > 0) sysChar = st[0];
                }

                // מיפוי לגרסת ה-SP הישנה:
                // - לכל מה שלא "תוכנית" → 'P'
                // - לתוכנית → 'ת'
                string spEntityType = MapEntityTypeForStoredProc(sysChar);

                // definement: לפי הגרסה הישנה – לתוכנית 0; לאחרות ערך ה-Definement
                int definement = (sysChar == 'ת') ? 0 : entity.Definement;

                // ניקוי מספר זיהוי (כמו קודם): לתוכנית לא מנקים; לאחרות מנקים
                string cleanedDsp = (sysChar == 'ת') ? (dspEntityNum ?? "") : CleanDspIfNeeded(dspEntityNum);

                using (var con = new SqlConnection(cs))
                using (var cmd = new SqlCommand("SP_Insert_Archive", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                    // חשוב: VarChar(1) כדי לאותת ל-SP את הקוד 'P' / 'ת' בדיוק כמו בישן
                    cmd.Parameters.Add("@entity_type", SqlDbType.VarChar, 1).Value = spEntityType;

                    cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = definement;

                    // נשאיר NVARCHAR כדי לתמוך בעברית בשדות טקסטואליים
                    cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.NVarChar, 100).Value = (object)(cleanedDsp ?? "") ?? "";
                    cmd.Parameters.Add("@File_Name", SqlDbType.NVarChar, 255).Value = (object)(fileDesc ?? "") ?? "";
                    cmd.Parameters.Add("@File_Location", SqlDbType.NVarChar, -1).Value = (object)(fullPath ?? "") ?? "";

                    // פרמטר החזרה + לוג InfoMessage
                    var ret = cmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int);
                    ret.Direction = ParameterDirection.ReturnValue;

                    con.InfoMessage += (s, e) => { try { DevDiag.Log("SQLMsg: " + e.Message); } catch { } };
                    con.FireInfoMessageEventOnUserErrors = true;

                    // לוג יעד
                    try
                    {
                        var b = new SqlConnectionStringBuilder(cs);
                        DevDiag.Log($"DBTargetCS: server={b.DataSource} db={b.InitialCatalog} user={b.UserID}");
                    }
                    catch { }

                    con.Open();
                    int rows = cmd.ExecuteNonQuery();

                    // לוג EXEC קריא
                    try
                    {
                        int rv = 0; try { rv = (ret.Value is int i) ? i : 0; } catch { }
                        DevDiag.Log("DB: SP return value = " + rv);
                        DevDiag.Log("EXEC " + DevDiag.AsExec(cmd) + " | rows=" + rows);
                    }
                    catch { }

                    // ב-SP מסוים rows=-1 תקין; נשתמש ב-ReturnValue==0 כ"עבר"
                    int retVal = 0; try { retVal = (ret.Value is int i) ? i : 0; } catch { }
                    return retVal == 0;
                }
            }
            catch (Exception ex)
            {
                try { DevDiag.Log("DB EX: " + ex.Message); } catch { }
                return false;
            }
        }

        // תאימות לאחור לפי שם ישות (אם הדיאלוג מחזיר string בלבד)
        public bool TryInsertRecord(string entityName, string dspEntityNum, string fullPath, string fileDesc)
        {
            var cs = GetConnectionString();
            if (string.IsNullOrWhiteSpace(cs)) return false;

            try
            {
                // נאתר סוג מערכת (עברית) להגדרת המיפוי
                string sysTypeStr = GetSystemEntityType(cs, entityName); // "ת"/"ב"/"פ"/...
                char sysChar = (!string.IsNullOrWhiteSpace(sysTypeStr)) ? sysTypeStr.Trim()[0] : 'ב';
                string spEntityType = MapEntityTypeForStoredProc(sysChar);

                int definement = (sysChar == 'ת')
                    ? 0
                    : GetDefinementEntityType(cs, entityName);

                string cleanedDsp = (sysChar == 'ת')
                    ? (dspEntityNum ?? "")
                    : CleanDspIfNeeded(dspEntityNum);

                using (var con = new SqlConnection(cs))
                using (var cmd = new SqlCommand("SP_Insert_Archive", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                    cmd.Parameters.Add("@entity_type", SqlDbType.VarChar, 1).Value = spEntityType;
                    cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = definement;

                    cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.NVarChar, 100).Value = (object)(cleanedDsp ?? "") ?? "";
                    cmd.Parameters.Add("@File_Name", SqlDbType.NVarChar, 255).Value = (object)(fileDesc ?? "") ?? "";
                    cmd.Parameters.Add("@File_Location", SqlDbType.NVarChar, -1).Value = (object)(fullPath ?? "") ?? "";

                    var ret = cmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int);
                    ret.Direction = ParameterDirection.ReturnValue;

                    con.InfoMessage += (s, e) => { try { DevDiag.Log("SQLMsg: " + e.Message); } catch { } };
                    con.FireInfoMessageEventOnUserErrors = true;

                    try
                    {
                        var b = new SqlConnectionStringBuilder(cs);
                        DevDiag.Log($"DBTargetCS: server={b.DataSource} db={b.InitialCatalog} user={b.UserID}");
                    }
                    catch { }

                    con.Open();
                    int rows = cmd.ExecuteNonQuery();

                    try
                    {
                        int rv = 0; try { rv = (ret.Value is int i) ? i : 0; } catch { }
                        DevDiag.Log("DB: SP return value = " + rv);
                        DevDiag.Log("EXEC " + DevDiag.AsExec(cmd) + " | rows=" + rows);
                    }
                    catch { }

                    int retVal = 0; try { retVal = (ret.Value is int i) ? i : 0; } catch { }
                    return retVal == 0;
                }
            }
            catch (Exception ex)
            {
                try { DevDiag.Log("DB EX(legacy): " + ex.Message); } catch { }
                return false;
            }
        }

        // ========= Helpers =========

        private static string MapEntityTypeForStoredProc(char sysChar)
        {
            // נאמן לגמרי לגרסה הישנה שלך:
            // בקשות/פיקוח/כללי/תביעה וכו' → 'P'
            // תוכנית → 'ת'
            return (sysChar == 'ת') ? "ת" : "P";
        }

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

        // ניקוי כמקודם (ללא שומרון)
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

                var parts = Regex.Split(dsp, @"\D+");
                foreach (var p in parts)
                {
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
                cmd.Parameters.AddWithValue("@name", entityName);
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
                cmd.Parameters.AddWithValue("@name", entityName);
                con.Open();
                object o = cmd.ExecuteScalar();
                if (o != null && o != DBNull.Value)
                {
                    int.TryParse(o.ToString(), out code);
                }
            }
            return code;
        }
    }
}
