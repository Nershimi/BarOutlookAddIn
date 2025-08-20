using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using BarOutlookAddIn.Properties;


namespace BarOutlookAddIn.App_Code
{
    class SystemEntity
    {
        string serverName = Properties.Settings.Default.ServerAddressNB;
        string DBName = Properties.Settings.Default.NBDBName;
        string connectionStr;

        string entityType, systemEntityType, dspEntityNum;
        DataTable EntityListTable;
        DataTable EntityTypeListTable;

        //Empty CTOR
        public SystemEntity()
        {

        }

        public DataTable entityTypeList()
        {
            EntityTypeListTable = new DataTable();
            EntityTypeListTable.Columns.Add("Code", typeof(string));
            EntityTypeListTable.Columns.Add("DescriptionCode", typeof(string));

            EntityTypeListTable.Rows.Add("ב", "בקשות");
            EntityTypeListTable.Rows.Add("פ", "פיקוח");
            EntityTypeListTable.Rows.Add("ת", "ישות תכנונית");
            EntityTypeListTable.Rows.Add("כ", "ישות כללית");
            EntityTypeListTable.Rows.Add("ע", "תביעה");

            return EntityTypeListTable;
        }

        //Get Entity List
        public DataTable entityList()
        {
            connectionStr = Properties.Settings.Default.ConnectionString;

            string entityQuery = "SELECT [entity_type] ,[entity_name] ,[system_entity_type] FROM [dbo].[Entity_Type_Control] where [IsEntityActive] = 1 and entity_type >= 10 and [entity_name] != 'תיק בניין' UNION " +
                "SELECT [definement_entity_type] as entity_type, [Description_Code] as entity_name,  [Code_Identification]as system_entity_type " +
                " FROM [dbo].[System_Entity] where Code_Identification = 'ת' and definement_entity_type > 0";

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                try
                {

                    using (SqlCommand cmd = new SqlCommand(entityQuery, conn))
                    {
                        conn.Open();

                        EntityListTable = new DataTable("EntityList");

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(EntityListTable);
                            conn.Close();
                            da.Dispose();
                        }
                    }

                }
                catch (Exception ex)
                {
                    using (EventLog eventLog = new EventLog("Application"))
                    {
                        eventLog.Source = "Bartech Outlook Archive addin";
                        eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                    }
                    DataTable dt = new DataTable("EntityListTable" + ex.Message);
                    return dt;
                }
                finally
                {
                    conn.Close();
                }
            }

            return EntityListTable;
        }

        //Check If Entity Exist
        public bool CheckIfExist(string _systemEntityType, string _entityType, string _dspEntityNum)
        {
            entityType = _entityType;

            systemEntityType = _systemEntityType;
            dspEntityNum = _dspEntityNum;

            //For Requests and Pikuach
            if (systemEntityType == "ב" || systemEntityType == "פ" || systemEntityType == "ע" || systemEntityType == "כ")
            {
                entityType = GetEntityType();

                bool isShomron = Properties.Settings.Default.IsShomron;

                if (!isShomron)
                {
                    string[] numbers = Regex.Split(dspEntityNum, @"\D+");
                    foreach (string value in numbers)
                    {
                        if (!string.IsNullOrEmpty(value))
                        {
                            int i = int.Parse(value);

                            if (i.ToString().Length > 1)
                            {
                                dspEntityNum = i.ToString();
                            }
                        }
                    }
                }

                if (CheckRequestAndPikuach())
                {
                    return true;
                }
            }

            // For City Plans
            if (systemEntityType == "ת")
            {
                entityType = GetTabaType();

                if (CheckTownPlan())
                {
                    return true;
                }
            }

            return false;
        }

        private bool CheckRequestAndPikuach()
        {
            bool result = false;
            connectionStr = Properties.Settings.Default.ConnectionString;
            bool isShomron = Properties.Settings.Default.IsShomron;

            string findAdministrativeEntity = "SELECT [appeal_system_number] ,[dsp_organizations_Number] ,[entity_type] ,[system_entity_type] FROM[dbo].[Administrative_Entity] where " +
                    " entity_type = @EntityType and system_entity_type = @SystemEntityType and ([dsp_organizations_Number] like @DspOrganizationsNumber or @DspOrganizationsNumber = '' " +
                    "OR REPLACE([dsp_organizations_Number] , char(254),'') like @DspOrganizationsNumber or sub_appeal_user_number + secondery_appeal_user_number + appeal_user_number + cast(zip_appeal_user_number as varchar(50)) " +
                    "like REPLACE(REPLACE(replace(@DspOrganizationsNumber,'/',''), char(254),''),'\','') or cast(zip_appeal_user_number as varchar(50))+appeal_user_number+secondery_appeal_user_number+ sub_appeal_user_number " +
                    "like REPLACE(REPLACE(replace(@DspOrganizationsNumber,'/',''), char(254),''),'\','') or (cast(zip_appeal_user_number as varchar(50))    like @DspOrganizationsNumber and UPPER(db_name())<>'BAR_SMR_NV'))";

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        conn.Open();
                        cmd.CommandText = findAdministrativeEntity;
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@DspOrganizationsNumber", "%" + dspEntityNum + "%");
                        cmd.Parameters.AddWithValue("@EntityType", entityType);
                        cmd.Parameters.AddWithValue("@SystemEntityType", systemEntityType);

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {

                            //Loop through all the rows, retrieving the columns you need.
                            while (dr.Read())
                            {

                            }
                            if (dr.HasRows)
                            {
                                return true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    using (EventLog eventLog = new EventLog("Application"))
                    {
                        eventLog.Source = "Bartech Outlook Archive addin";
                        eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                    }
                    throw;
                }
                finally
                {
                    conn.Close();
                }
                return result;
            }
        }

        private bool CheckTownPlan()
        {
            bool result = false;
            connectionStr = Properties.Settings.Default.ConnectionString;

            string findAdministrativeEntity = "SELECT [Taba] ,[Description_Taba] FROM [dbo].[Types_TownPlan] where Taba = @Taba";

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        conn.Open();
                        cmd.CommandText = findAdministrativeEntity;
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@Taba", dspEntityNum);

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {

                            //Loop through all the rows, retrieving the columns you need.
                            while (dr.Read())
                            {

                            }
                            if (dr.HasRows)
                            {
                                return true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    using (EventLog eventLog = new EventLog("Application"))
                    {
                        eventLog.Source = "Bartech Outlook Archive addin";
                        eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                    }
                    throw;
                }
                finally
                {
                    conn.Close();
                }
                return result;
            }
        }

        private string GetEntityType()
        {
            string result = "";

            connectionStr = Properties.Settings.Default.ConnectionString;

            string getEntityTypeQuery = "SELECT  [entity_type]  FROM [dbo].[Entity_Type_Control] where [entity_name] = @EntityName";

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                using (SqlCommand comm = new SqlCommand(getEntityTypeQuery))
                {
                    comm.Connection = conn;
                    comm.Parameters.AddWithValue("@EntityName", entityType);

                    try
                    {
                        conn.Open();
                        DataTable dt = new DataTable();
                        dt.Load(comm.ExecuteReader());

                        foreach (DataRow row in dt.Rows)
                        {
                            result = row[0].ToString();
                        }
                        return result;

                    }
                    catch (SqlException ex)
                    {
                        using (EventLog eventLog = new EventLog("Application"))
                        {
                            eventLog.Source = "Bartech Outlook Archive addin";
                            eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                        }
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        conn.Close();
                    }
                }
                return result;
            }
        }

        private string GetTabaType()
        {
            string result = "";

            connectionStr = Properties.Settings.Default.ConnectionString;

            string getEntityTypeQuery = "SELECT  [planning_information_type]  FROM [dbo].[Types_TownPlan] where [Taba] = @Taba";

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                using (SqlCommand comm = new SqlCommand(getEntityTypeQuery))
                {
                    comm.Connection = conn;
                    comm.Parameters.AddWithValue("@Taba", dspEntityNum);

                    try
                    {
                        conn.Open();
                        DataTable dt = new DataTable();
                        dt.Load(comm.ExecuteReader());

                        foreach (DataRow row in dt.Rows)
                        {
                            result = row[0].ToString();
                        }
                        return result;

                    }
                    catch (SqlException ex)
                    {
                        using (EventLog eventLog = new EventLog("Application"))
                        {
                            eventLog.Source = "Bartech Outlook Archive addin";
                            eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                        }
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        conn.Close();
                    }
                }
                return result;
            }
        }
        public List<string> GetEntityNames()
        {
            List<string> entities = new List<string>();
            DataTable dt = this.entityList();

            foreach (DataRow row in dt.Rows)
            {
                if (dt.Columns.Contains("entity_name"))
                {
                    string name = row["entity_name"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        entities.Add(name);
                    }
                }
            }

            return entities;
        }


        #region Sql and DB Write

        //Check If DB Exist
        public bool checkIfDBExist()
        {
            bool result = false;

            connectionStr = Properties.Settings.Default.ConnectionString;

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(connectionStr))
                {
                    sqlConnection.Open();
                    sqlConnection.Close();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Bartech Outlook Archive addin";
                    eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                }
                System.Windows.Forms.MessageBox.Show(ex.Message);
                result = false;
            }
            return result;
        }

        //Sql Write
        public bool writeArchiveToDB(string _systemEntityType, string _entityType, string _dspEntityNum, string _fileDesc, string _filePath)
        {
            bool isShomron = Properties.Settings.Default.IsShomron;
            bool result = false;
            connectionStr = Properties.Settings.Default.ConnectionString;
            entityType = _entityType;
            entityType = GetEntityType();
            systemEntityType = _systemEntityType;
            dspEntityNum = _dspEntityNum;

            string fileDesc = _fileDesc;
            string filePath = _filePath;

            //For Requests and Pikuach
            if (systemEntityType == "ב" || systemEntityType == "פ" || systemEntityType == "ע" || systemEntityType == "כ")
            {
                if (!isShomron)
                {
                    string[] numbers = Regex.Split(dspEntityNum, @"\D+");
                    foreach (string value in numbers)
                    {
                        if (!string.IsNullOrEmpty(value))
                        {
                            int i = int.Parse(value);

                            if (i.ToString().Length > 1)
                            {
                                dspEntityNum = i.ToString();
                            }
                        }
                    }
                }

                if (CheckRequestAndPikuach())
                {
                    int entityTypeInt = Convert.ToInt32(entityType);
                    string dspEntityNumInt = dspEntityNum;

                    using (SqlConnection con = new SqlConnection(connectionStr))
                    {
                        try
                        {
                            using (SqlCommand cmd = new SqlCommand("SP_Insert_Archive", con))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                                cmd.Parameters.Add("@entity_type", SqlDbType.VarChar).Value = 'P';
                                cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = entityTypeInt;
                                cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.VarChar).Value = dspEntityNumInt;
                                cmd.Parameters.Add("@File_Name", SqlDbType.VarChar).Value = fileDesc;
                                cmd.Parameters.Add("@File_Location", SqlDbType.VarChar).Value = filePath;

                                con.Open();
                                cmd.ExecuteNonQuery();
                                result = true;
                            }

                        }
                        catch (Exception ex)
                        {
                            using (EventLog eventLog = new EventLog("Application"))
                            {
                                eventLog.Source = "Bartech Outlook Archive addin";
                                eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                            }
                            result = false;
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            con.Close();
                        }
                    }

                    return result;
                }
            }
            // For City Plans
            if (systemEntityType == "ת")
            {
                if (CheckTownPlan())
                {
                    using (SqlConnection con = new SqlConnection(connectionStr))
                    {
                        try
                        {
                            using (SqlCommand cmd = new SqlCommand("SP_Insert_Archive", con))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.Add("@Estate_Number", SqlDbType.BigInt).Value = 0;
                                cmd.Parameters.Add("@entity_type", SqlDbType.VarChar).Value = 'ת';
                                cmd.Parameters.Add("@definement_entity_type", SqlDbType.Int).Value = 0;
                                cmd.Parameters.Add("@Org_Entity_Number", SqlDbType.VarChar).Value = dspEntityNum;
                                cmd.Parameters.Add("@File_Name", SqlDbType.VarChar).Value = fileDesc;
                                cmd.Parameters.Add("@File_Location", SqlDbType.VarChar).Value = filePath;

                                con.Open();
                                cmd.ExecuteNonQuery();
                                result = true;
                            }

                        }
                        catch (Exception ex)
                        {
                            using (EventLog eventLog = new EventLog("Application"))
                            {
                                eventLog.Source = "Bartech Outlook Archive addin";
                                eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                            }
                            result = false;
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            con.Close();
                        }
                    }
                }
                else
                {
                    return result;
                }
            }

            if (systemEntityType == "u")
            {
                return result;
            }
            else
            {
                return result;
            }
        }
    }
}

#endregion



