using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace BarOutlookAddIn.Helpers
{
    public static class DbInspector
    {
        // Return true if the stored procedure exists in the target database (schema.name form, e.g. "dbo.SP_Insert_Archive")
        public static bool StoredProcedureExists(string connectionString, string fullProcName)
        {
            if (string.IsNullOrWhiteSpace(connectionString)) throw new ArgumentNullException(nameof(connectionString));
            if (string.IsNullOrWhiteSpace(fullProcName)) throw new ArgumentNullException(nameof(fullProcName));

            var parts = fullProcName.Split(new[] {'.'}, 2);
            string schema = parts.Length == 2 ? parts[0] : "dbo";
            string name = parts.Length == 2 ? parts[1] : parts[0];

            using (var conn = new SqlConnection(connectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    @"SELECT COUNT(1)
                      FROM sys.procedures p
                      WHERE p.name = @name AND SCHEMA_NAME(p.schema_id) = @schema;";
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@schema", schema);

                conn.Open();
                var result = cmd.ExecuteScalar();
                return Convert.ToInt32(result) > 0;
            }
        }

        // Returns list of parameters as "name type(length/nullability)" in the order defined in the DB.
        public static List<string> GetStoredProcedureParameters(string connectionString, string fullProcName)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(connectionString)) return list;
            if (string.IsNullOrWhiteSpace(fullProcName)) return list;

            var parts = fullProcName.Split(new[] {'.'}, 2);
            string schema = parts.Length == 2 ? parts[0] : "dbo";
            string name = parts.Length == 2 ? parts[1] : parts[0];

            const string sql = @"
SELECT
    p.PARAMETER_NAME,
    p.DATA_TYPE,
    p.CHARACTER_MAXIMUM_LENGTH,
    p.PARAMETER_MODE
FROM INFORMATION_SCHEMA.PARAMETERS p
WHERE p.SPECIFIC_SCHEMA = @schema AND p.SPECIFIC_NAME = @name
ORDER BY p.ORDINAL_POSITION;";

            using (var conn = new SqlConnection(connectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue("@schema", schema);
                cmd.Parameters.AddWithValue("@name", name);

                conn.Open();
                using (var rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        var pname = rdr["PARAMETER_NAME"] as string;
                        var dtype = rdr["DATA_TYPE"] as string;
                        var len = rdr["CHARACTER_MAXIMUM_LENGTH"] is DBNull ? null : (rdr["CHARACTER_MAXIMUM_LENGTH"] as int?);
                        var mode = rdr["PARAMETER_MODE"] as string;

                        string lenPart = len.HasValue && len.Value > 0 ? "(" + len.Value + ")" : "";
                        list.Add($"{pname} {dtype}{lenPart} {mode}");
                    }
                }
            }

            return list;
        }

        // Quick comparer: pass expected list like "@Estate_Number bigint", "@entity_type char(1)", ...
        public static bool MatchesExpectedSignature(string connectionString, string fullProcName, IList<string> expectedParameters, out string message)
        {
            message = null;
            try
            {
                if (!StoredProcedureExists(connectionString, fullProcName))
                {
                    message = "Procedure not found.";
                    return false;
                }

                var actual = GetStoredProcedureParameters(connectionString, fullProcName);
                if (actual.Count != expectedParameters.Count)
                {
                    message = $"Parameter count mismatch. expected={expectedParameters.Count}, actual={actual.Count}";
                    return false;
                }

                for (int i = 0; i < expectedParameters.Count; i++)
                {
                    // compare normalized strings (case-insensitive, trim)
                    if (!string.Equals(NormalizeParam(expectedParameters[i]), NormalizeParam(actual[i]), StringComparison.OrdinalIgnoreCase))
                    {
                        message = $"Parameter mismatch at position {i+1}: expected='{expectedParameters[i]}', actual='{actual[i]}'";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                message = "Error: " + ex.Message;
                return false;
            }
        }

        private static string NormalizeParam(string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return "";
            return p.Replace(" ", "").Replace("\t", "").Replace("(", "").Replace(")", "").Replace("output", "").Replace("INPUT", "").Trim().ToLowerInvariant();
        }
    }
}