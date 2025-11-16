using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

namespace BarOutlookAddIn
{
    /// <summary>
    /// Service for an atomic running number used by the archive. Matches table dbo.Numerators (Number_Numerator, Numerator).
    /// Not dependent on a stored procedure; uses a transaction and locking hints to avoid conflicts.
    /// Adds retry on transient deadlocks to improve robustness.
    /// </summary>
    internal static class NumeratorService
    {
        /// <summary>
        /// Default code for the archive numerator (same as the old add-in: Number_Numerator = 5).
        /// </summary>
        public const int ArchiveNumberCode = 5;

        /// <summary>
        /// Returns the next value atomically for a given Number_Numerator.
        /// Retries on deadlock to reduce transient failures.
        /// </summary>
        public static int GetNext(int numberCode)
        {
            string cs = GetConnectionString();
            const int maxAttempts = 3;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    using (var conn = new SqlConnection(cs))
                    {
                        conn.Open();
                        using (var tx = conn.BeginTransaction(IsolationLevel.Serializable))
                        {
                            // Try to update an existing row and return the new value using OUTPUT
                            object updResult;
                            using (var cmdUpdate = new SqlCommand(@"
UPDATE dbo.Numerators WITH (UPDLOCK, HOLDLOCK)
SET Numerator = Numerator + 1
OUTPUT inserted.Numerator
WHERE Number_Numerator = @code;", conn, tx))
                            {
                                cmdUpdate.Parameters.Add(new SqlParameter("@code", SqlDbType.Int) { Value = numberCode });
                                // ExecuteScalar returns first column of first row or null when no rows returned
                                updResult = cmdUpdate.ExecuteScalar();
                            }

                            int nextVal;
                            // If update affected a row, we got the new numerator
                            if (updResult != null && updResult != DBNull.Value)
                            {
                                nextVal = Convert.ToInt32(updResult);
                            }
                            else
                            {
                                // No row existed — insert a new one with value 1
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
                            return nextVal;
                        }
                    }
                }
                catch (SqlException sqlEx) when (sqlEx.Number == 1205 || sqlEx.Number == 1222) // deadlock / lock request timeout
                {
                    if (attempt == maxAttempts)
                        throw;
                    // Back off a bit and retry
                    Thread.Sleep(100 * attempt);
                    continue;
                }
            }

            throw new InvalidOperationException("Failed to allocate next numerator after retries.");
        }

        /// <summary>
        /// The next value for the archive numerator (Number_Numerator = 5).
        /// </summary>
        public static int GetNextArchiveNumber() => GetNext(ArchiveNumberCode);

        private static string GetConnectionString()
        {
            // Explicit read from settings (consistent with the rest of the code)
            var settings = global::BarOutlookAddIn.Properties.Settings.Default;
            var props = settings.Properties;
            if (props == null || props["ConnectionString"] == null)
                throw new InvalidOperationException("ConnectionString is not defined in Settings.");

            var cs = settings["ConnectionString"] as string;
            if (string.IsNullOrWhiteSpace(cs))
                throw new InvalidOperationException("ConnectionString is empty or invalid.");
            return cs;
        }
    }
}
