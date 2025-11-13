using System;
using System.Data;
using System.Data.SqlClient;

namespace BarOutlookAddIn
{
    /// <summary>
    /// שירות למספר רץ (אטומי) עבור ארכיון. תואם לטבלת dbo.Numerators (Number_Numerator, Numerator).
    /// לא תלוי ב-SP; עובד עם טרנזקציה ו-Hints למניעת התנגשויות.
    /// </summary>
    internal static class NumeratorService
    {
        /// <summary>
        /// קוד ברירת המחדל למונה הארכיון (כמו בתוסף הישן: Number_Numerator = 5).
        /// </summary>
        public const int ArchiveNumberCode = 5;

        /// <summary>
        /// מחזיר את הערך הבא בצורה אטומית עבור Number_Numerator נתון.
        /// </summary>
        public static int GetNext(int numberCode)
        {
            string cs = GetConnectionString();
            using (var conn = new SqlConnection(cs))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction(IsolationLevel.Serializable))
                {
                    // ננסה לעדכן קיים ולהחזיר את הערך החדש בעזרת OUTPUT
                    int? nextVal = null;
                    using (var cmdUpdate = new SqlCommand(@"
UPDATE dbo.Numerators WITH (UPDLOCK, HOLDLOCK)
SET Numerator = Numerator + 1
OUTPUT inserted.Numerator
WHERE Number_Numerator = @code;", conn, tx))
                    {
                        cmdUpdate.Parameters.Add(new SqlParameter("@code", SqlDbType.Int) { Value = numberCode });

                        using (var rdr = cmdUpdate.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                nextVal = rdr.GetInt32(0);
                            }
                        }
                    }

                    // אם לא היתה שורה – ניצור חדשה עם 1
                    if (!nextVal.HasValue)
                    {
                        using (var cmdInsert = new SqlCommand(@"
INSERT INTO dbo.Numerators(Number_Numerator, Numerator)
VALUES(@code, 1);
SELECT CAST(1 AS INT);", conn, tx))
                        {
                            cmdInsert.Parameters.Add(new SqlParameter("@code", SqlDbType.Int) { Value = numberCode });
                            nextVal = Convert.ToInt32(cmdInsert.ExecuteScalar());
                        }
                    }

                    tx.Commit();
                    return nextVal.Value;
                }
            }
        }

        /// <summary>
        /// הערך הבא למונה הארכיון (Number_Numerator=5).
        /// </summary>
        public static int GetNextArchiveNumber() => GetNext(ArchiveNumberCode);

        private static string GetConnectionString()
        {
            // קריאה מפורשת להגדרות (כמו בשאר הקוד שלך)
            var settings = global::BarOutlookAddIn.Properties.Settings.Default;
            var props = settings.Properties;
            if (props == null || props["ConnectionString"] == null)
                throw new InvalidOperationException("ConnectionString לא מוגדר ב-Settings.");

            var cs = settings["ConnectionString"] as string;
            if (string.IsNullOrWhiteSpace(cs))
                throw new InvalidOperationException("ConnectionString ריק או לא תקף.");
            return cs;
        }
    }
}
