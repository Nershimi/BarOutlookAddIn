using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace BarOutlookAddIn.Helpers
{
    public class EntityRepository
    {
        private readonly string _connectionString;

        public EntityRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // Load entities with the same logic as the legacy add-in.
        public List<EntityInfo> GetEntities()
        {
            var list = new List<EntityInfo>();

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT
    CAST(entity_type AS INT)          AS Definement,
    entity_name                       AS Name,
    system_entity_type                AS SystemType
FROM dbo.Entity_Type_Control
WHERE IsEntityActive = 1
  AND entity_type >= 10
  AND entity_name <> N'תיק בניין'
UNION ALL
SELECT
    CAST(definement_entity_type AS INT) AS Definement,
    Description_Code                     AS Name,
    Code_Identification                  AS SystemType
FROM dbo.System_Entity
WHERE Code_Identification = N'ת'
  AND definement_entity_type > 0
ORDER BY Name;";

                conn.Open();
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var item = new EntityInfo();
                        object o;

                        o = r["Definement"];
                        item.Definement = o != null && o != System.DBNull.Value ? Convert.ToInt32(o) : 0;

                        item.Name = r["Name"] as string;
                        item.SystemType = r["SystemType"] as string;

                        if (!string.IsNullOrWhiteSpace(item.Name))
                            list.Add(item);
                    }
                }
            }

            return list;
        }
    }
}
