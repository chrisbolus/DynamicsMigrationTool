using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Deployment;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using XrmToolBox.Extensibility;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;
using View = Microsoft.SqlServer.Management.Smo.View;
using Server = Microsoft.SqlServer.Management.Smo.Server;
using System.Data.SqlClient;
using System.Activities.Expressions;
using System.Diagnostics.Metrics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace DynamicsMigrationTool
{
    public partial class MyPluginControl : PluginControlBase
    {
        private Settings mySettings;

        public MyPluginControl()
        {
            InitializeComponent();
        }

        private class EntityAttribute_AdditionalInfo
        {
            public string DBDataType;
            public string SSISDataType;
            public int? StringLength = null;
            public bool isValidForMigration = true;
            public bool isLookup = false;
            public bool isPrimaryKey = false;
            public string DynEntityLookupTargets_List;
            public string DynDataType_Readable;
        }


        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");

                if(Service != null)
                {
                    SetEntityList();
                }

                sourceDBConnection_txtb.Text = mySettings.SourceDBConnectionString;
                stagingDBConnection_txtb.Text = mySettings.StagingDBConnectionString;

            }
        }

        private void SetEntityList()
        {
            var EntityList = GetEntities(Service);

            EntityCmb.DataSource = EntityList;
            EntityCmb.DisplayMember = "LogicalName";                
            EntityCmb.SelectedIndex = -1;       //Default EntityCmb to blank
        }


        public EntityMetadata[] GetEntities(IOrganizationService organizationService)
        {
            Dictionary<string, string> attributesData = new Dictionary<string, string>();
            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
            metaDataRequest.EntityFilters = EntityFilters.Entity;



            XmlDictionaryReaderQuotas myReaderQuotas = new XmlDictionaryReaderQuotas();
            myReaderQuotas.MaxNameTableCharCount = 2147483647;



            // Execute the request.



            metaDataResponse = (RetrieveAllEntitiesResponse)organizationService.Execute(metaDataRequest);



            var entities = metaDataResponse.EntityMetadata;
            return entities;
        }


        private void GetEntityMetaData()
        {
            GetEntities(Service);

        }

        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
            SetEntityList();
        }

        private void CreateStgTbl_Btn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(TestConnection);

            if (IsEntitySelected())
            {
                var entityMetadata_noattr = (EntityMetadata)EntityCmb.SelectedItem;

                var entityMetadata = Service.GetEntityMetadata(entityMetadata_noattr.LogicalName);

                var stagingDBConnectionString = new SqlConnection();

                Boolean isStagingDBConnectionValid = false;
                try
                {
                    stagingDBConnectionString = new SqlConnection(mySettings.StagingDBConnectionString);
                    stagingDBConnectionString.Open();
                    isStagingDBConnectionValid = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Please review Staging Database Connection String:\n{mySettings.StagingDBConnectionString}\n\nError: {ex.Message}\n\nExample Connection String:\nData Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Staging_DB;Integrated Security=True;");
                }

                if (isStagingDBConnectionValid)
                {
                    CreateStagingTable(stagingDBConnectionString, entityMetadata);

                    MessageBox.Show("Staging Table Created Successfully");
                }
            }
        }


        private void CreateSrcVwTmpl_Btn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(TestConnection);

            if (IsEntitySelected())
            {
                var entityMetadata_noattr = (EntityMetadata)EntityCmb.SelectedItem;

                var entityMetadata = Service.GetEntityMetadata(entityMetadata_noattr.LogicalName);

                var sourceDBConnectionString = new SqlConnection();

                Boolean isSourceDBConnectionValid = false;
                try
                {
                    sourceDBConnectionString = new SqlConnection(mySettings.SourceDBConnectionString);
                    sourceDBConnectionString.Open(); 
                    isSourceDBConnectionValid = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Please review Source Database Connection String:\n{mySettings.SourceDBConnectionString}\n\nError: {ex.Message}\n\nExample Connection String:\nData Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Source_DB;Integrated Security=True;");
                }


                Boolean doesDMTSchemaExist = false;

                if (isSourceDBConnectionValid)
                {

                    var getBadRecordsStg = new SqlCommand("select schema_id from sys.schemas where name = 'DMT'", sourceDBConnectionString);

                    var schemaCount = 0;

                    using (SqlDataReader rdr = getBadRecordsStg.ExecuteReader())
                    {

                        while (rdr.Read())
                        {
                            schemaCount = 1;
                        }
                    }
                    if (schemaCount > 0)
                    {
                        doesDMTSchemaExist = true;
                    }
                    if (schemaCount == 0)
                    {

                        var result = MessageBox.Show("The Source Database requires the a schema called \"DMT\" which doesn't exist. Add schema to source database?", "Add DMT Schema?",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            try
                            {

                                new SqlCommand("CREATE SCHEMA [DMT]", sourceDBConnectionString).ExecuteNonQuery();
                                doesDMTSchemaExist = true;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Failed to create schema in Source Database. " + ex.Message);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Unable to proceed without DMT schema.");
                        }


                    }

                }

                if (doesDMTSchemaExist)
                {
                    CreateTemplateSourceView(sourceDBConnectionString, entityMetadata);

                    MessageBox.Show("Source View Template Created Successfully");
                }
            }
        }

        private void TestConnection()
        {
            var a = 1;
        }


        private bool IsEntitySelected()
        {
            if (EntityCmb.SelectedItem == null)
            {
                if (EntityCmb.Text == null)
                {
                    MessageBox.Show("Please select an entity from the dropdown.");
                    return false;
                }
                else if (EntityCmb.Text != null)
                {
                    //var entities = GetEntities(Service);
                    //foreach (var entity in entities)
                    //{
                    //    if (entity.LogicalName == EntityCmb.Text)
                    //    {
                    //        EntityCmb.SelectedItem = entity;
                    //        return true;
                    //    }
                    //}

                    MessageBox.Show("Please select an entity from the dropdown.");
                    return false;
                }
            }
            return true;
        }

        void OnPropertyChanged(String prop)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<EntityMetadata> entityList;

        public ObservableCollection<EntityMetadata> EntityList
        {
            get { return entityList; }
            set
            {
                entityList = value;
                OnPropertyChanged("EntityList");
            }
        }


        public void CreateTemplateSourceView(SqlConnection connection, EntityMetadata entity)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandTimeout = 0;
                command.CommandText = $@"
CREATE OR ALTER VIEW [dmt].{entity.LogicalName}_Template AS 
--RENAME THIS VIEW! Remove the ""_Template\"" suffix, or replace it with your own suffix.
--THIS IS A BASE VIEW OF ALL FIELDS FOR THE {entity.LogicalName.ToUpper()} ENTITY WITH ALL THE CASTS TO ENSURE ALL THE DATA ENDS UP IN THE CORRECT FORMAT. ADD A FROM STATEMENT TO PULL DATA FROM THE TABLE YOU WANT TO USE, AND REPLACE THE NULLS IN THE CASTS WITH THE FIELDS FROM YOUR SOURCE TABLE .

SELECT
";

                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(1 AS INT) AS Source_System_Id", 
                                                                        $"DMT Field, used to handle multiple source systems. Default value = 1. Change if using a second (etc.) source system.", true);
                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(1 AS INT) AS Processing_Status", 
                                                                        $"DMT Field, used to handle relationship mapping. Default value = 1. Set to 0 for records that shouldn't be loaded.");


                foreach (var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
                {
                    EntityAttribute_AdditionalInfo entAAI = Get_EntityAttribute_AdditionalInfo(attribute);


                    if (entAAI.isValidForMigration)
                    {
                        var targetListMessage = entAAI.DynEntityLookupTargets_List == null ? "" : $" Target entity(s) = {entAAI.DynEntityLookupTargets_List}.";

                        command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS {entAAI.DBDataType}) AS {attribute.LogicalName}", 
                                                                                $"Dynamics Attribute Type = {entAAI.DynDataType_Readable}.{targetListMessage}");
                        if (entAAI.isLookup)
                        {
                            if (entAAI.isPrimaryKey)
                            {
                                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Source",
                                                                                        $"DMT Field, used to handle the id of the record from the source system. Staging table requires this to be populated and unique when combined with Source_System_Id.");
                                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Leader",
                                                                                        $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with Leader_System_Id.");
                                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS INT) AS Leader_System_Id",
                                                                                        $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with {attribute.LogicalName}_Leader.");
                            }
                            else
                            {
                                command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Source",
                                                                                        $"DMT Field, used to handle the source system ids of other entities.{targetListMessage}");
                            }
                        }
                    }
                }

                command.CommandText += "\n--FROM ";

                command.ExecuteNonQuery();
            }

        }

        public string TemplateSourceView_HandleSpacing(string queryLine, string comment, bool isFirstLine = false)
        {
            int spacing = 80;
            int spacingoffset = queryLine.Length >= spacing ? 10 : spacing - queryLine.Length;
            string addComma = isFirstLine ? " " : ",";

            return $"       {addComma}{queryLine}{new String(' ', spacingoffset)}--{comment}\n";
        }


        public void CreateStagingTable(SqlConnection connection, EntityMetadata entity)
        {
            var dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");

            using (var command = connection.CreateCommand())
            {
                command.CommandTimeout = 0;
                command.CommandText = $@"
                    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].{entity.LogicalName}') AND type in (N'U'))
                    DROP TABLE [dbo].{entity.LogicalName}

                    CREATE TABLE [dbo].{entity.LogicalName} ( 
                    Global_Id INT IDENTITY(1,1), 
                    Source_System_Id INT NOT NULL, 
                    Processing_Status INT NOT NULL, 
                    FlagCreate BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagCreate_{dateTimeNow} DEFAULT 1, 
                    FlagUpdate BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagUpdate_{dateTimeNow} DEFAULT 0, 
                    FlagDelete BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagDelete_{dateTimeNow} DEFAULT 0, 
                    DateCreate DATETIME, 
                    DateUpdate DATETIME, 
                    DateDelete DATETIME, 
                    DynCreateId UNIQUEIDENTIFIER, 
                    DynUpdateId UNIQUEIDENTIFIER, 
                    DynDeleteId UNIQUEIDENTIFIER 
                    );" + "\n\n";



                foreach (var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
                {
                    EntityAttribute_AdditionalInfo entAAI = Get_EntityAttribute_AdditionalInfo(attribute);

                    if (entAAI.isValidForMigration)
                    {
                        command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD {attribute.LogicalName} {entAAI.DBDataType};\n";

                        if (entAAI.isLookup)
                        {
                            if (entAAI.isPrimaryKey)
                            {
                                command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD {attribute.LogicalName}_Source NVARCHAR(255) NOT NULL;\n";
                                command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD {attribute.LogicalName}_Leader NVARCHAR(255);\n";
                                command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD Leader_System_Id INT;\n";

                            }
                            else
                            {
                                command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD {attribute.LogicalName}_Source NVARCHAR(255);\n";
                            }
                        }
                    }


                }
                command.ExecuteNonQuery();
            }

            CreateIndexes(connection, entity);
            CreateRelationshipMapView(connection, entity);
        }

        private EntityAttribute_AdditionalInfo Get_EntityAttribute_AdditionalInfo(AttributeMetadata attribute)
        {
            var EntAAI = new EntityAttribute_AdditionalInfo();
            EntAAI.DynDataType_Readable = attribute.AttributeType?.ToString();      //this is modified in a few of cases below.

            if ((attribute.AttributeOf != null && attribute.AttributeType == AttributeTypeCode.String)
                || (attribute.IsValidForCreate == false && attribute.IsValidForUpdate == false)
                || (attribute.IsLogical == true && attribute.IsPrimaryId == true)
                || attribute.AttributeType == AttributeTypeCode.PartyList
                || attribute.AttributeType == AttributeTypeCode.ManagedProperty)
            {
                //Bad Records
                EntAAI.isValidForMigration = false;
            }
            else if (attribute.AttributeType == AttributeTypeCode.BigInt)
            {
                EntAAI.DBDataType = "BIGINT";
                EntAAI.SSISDataType = "i8";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Boolean
                || attribute.AttributeType == AttributeTypeCode.EntityName
                || attribute.AttributeType == AttributeTypeCode.Picklist
                || attribute.AttributeType == AttributeTypeCode.State
                || attribute.AttributeType == AttributeTypeCode.Status)
            {
                EntAAI.DBDataType = "NVARCHAR(255)";
                EntAAI.SSISDataType = "wstr";
                EntAAI.StringLength = 255;
            }
            else if (attribute.AttributeType == AttributeTypeCode.Customer
                || attribute.AttributeType == AttributeTypeCode.Lookup
                || attribute.AttributeType == AttributeTypeCode.Owner
                || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier)
            {
                EntAAI.DBDataType = "UNIQUEIDENTIFIER";
                EntAAI.SSISDataType = "guid";
                EntAAI.isLookup = true;
                EntAAI.DynEntityLookupTargets_List = GetEntityLookupTargetsList(attribute);

                if (attribute.AttributeType == AttributeTypeCode.Uniqueidentifier)
                {
                    var entityMetadata = Service.GetEntityMetadata(attribute.EntityLogicalName);
                    if (entityMetadata.PrimaryIdAttribute == attribute.LogicalName)
                    {
                        EntAAI.isPrimaryKey = true;
                    }
                }
            }
            else if (attribute.AttributeType == AttributeTypeCode.DateTime)
            {
                EntAAI.DBDataType = "DATETIME";
                EntAAI.SSISDataType = "dbTimeStamp";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Decimal)
            {
                EntAAI.DBDataType = "DECIMAL(23, 10)";
                EntAAI.SSISDataType = "numeric";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Double)
            {
                EntAAI.DBDataType = "FLOAT";
                EntAAI.SSISDataType = "r8";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Integer)
            {
                EntAAI.DBDataType = "INT";
                EntAAI.SSISDataType = "i4";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Memo)
            {
                EntAAI.DBDataType = "NVARCHAR(MAX)";
                EntAAI.SSISDataType = "nText";
                EntAAI.StringLength = -1;
            }
            else if (attribute.AttributeType == AttributeTypeCode.Money)
            {
                EntAAI.DBDataType = "MONEY";
                EntAAI.SSISDataType = "cy";
            }
            else if (attribute.AttributeType == AttributeTypeCode.String)
            {
                var strlength = (int)attribute.GetType().GetProperty("MaxLength").GetValue(attribute);
                EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + $"({strlength})";
                if (strlength > 4000)
                {
                    EntAAI.DBDataType = "NVARCHAR(MAX)";
                    EntAAI.SSISDataType = "nText";
                    EntAAI.StringLength = -1;
                }
                else
                {
                    EntAAI.DBDataType = $"NVARCHAR({strlength})";
                    EntAAI.SSISDataType = "wstr";
                    EntAAI.StringLength = strlength;
                }
            }
            else if (attribute.AttributeType == AttributeTypeCode.Virtual)
            {
                if (attribute.AttributeTypeName.Value == "MultiSelectPicklistType")
                {
                    EntAAI.DBDataType = "NVARCHAR(1000)";
                    EntAAI.SSISDataType = "wstr";
                    EntAAI.StringLength = 1000;
                    EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + " - " + attribute.AttributeTypeName.Value;
                }
                else if (attribute.AttributeTypeName.Value == "ImageType")
                {
                    EntAAI.DBDataType = "IMAGE";
                    EntAAI.SSISDataType = "image";
                    EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + " - " + attribute.AttributeTypeName.Value;
                }
                else
                {
                    //attribute.AttributeTypeName.Value == "VirtualType"
                    EntAAI.isValidForMigration = false;
                }
            }
            else
            {
                //In theory this should never be hit.
                EntAAI.isValidForMigration = false;
            }

            return EntAAI;

        }

        private string GetEntityLookupTargetsList(AttributeMetadata attribute)
        {
            if(attribute.GetType().GetProperty("Targets") != null)
            {
                var targets = (string[])attribute.GetType().GetProperty("Targets").GetValue(attribute);

                return String.Join(",", targets);
            }
            else return null;
        }

        private void CreateIndexes(SqlConnection connection, EntityMetadata entity)
        {
            var dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");

            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT PK_{entity.LogicalName}_{dateTimeNow} PRIMARY KEY CLUSTERED (Source_System_Id ASC, {entity.PrimaryIdAttribute}_Source ASC) ON [PRIMARY]
                CREATE INDEX IDX_Processing_Status ON dbo.{entity.LogicalName}(Processing_Status, FlagCreate, DynCreateId, FlagUpdate, DynUpdateId, FlagDelete, DynDeleteId);
                CREATE INDEX IDX_Global_Id ON dbo.{entity.LogicalName}(Global_Id);
                CREATE INDEX IDX_CreateParameters ON dbo.{entity.LogicalName}(FlagCreate, DynCreateId);
                CREATE INDEX IDX_UpdateParameters ON dbo.{entity.LogicalName}(FlagUpdate, DynUpdateId);
                CREATE INDEX IDX_DeleteParameters ON dbo.{entity.LogicalName}(FlagDelete, DynDeleteId);
        ";
                command.ExecuteNonQuery();
            }
        }

        private void CreateRelationshipMapView(SqlConnection connection, EntityMetadata entity)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
                CREATE OR ALTER VIEW dbo.{entity.LogicalName}_RelationshipMap AS
                SELECT
                    Follower.Source_System_Id,
                    Follower.{entity.PrimaryIdAttribute}_Source,
                    Follower.Leader_System_Id,
                    Follower.{entity.PrimaryIdAttribute}_Leader,
                    Follower.Processing_Status,
                    COALESCE(Leader.DynCreateId, Follower.DynCreateId) AS DynCreateId
                FROM dbo.{entity.LogicalName} AS Follower
                LEFT JOIN dbo.{entity.LogicalName} AS Leader
                        ON Follower.Leader_System_Id = Leader.Source_System_Id AND Follower.{entity.PrimaryIdAttribute}_Leader = Leader.{entity.PrimaryIdAttribute}_Source;
                ";
                command.ExecuteNonQuery();
            }
        }

        private bool attributeToExclude(string attribute_LogicalName)
        {
            if (attribute_LogicalName != null)
            { 
                if (attribute_LogicalName == "createdby"
                    || attribute_LogicalName == "createdbyexternalparty"
                    || attribute_LogicalName == "createdon"
                    || attribute_LogicalName == "createdonbehalfby"
                    || attribute_LogicalName == "modifiedby"
                    || attribute_LogicalName == "modifiedbyexternalparty"
                    || attribute_LogicalName == "modifiedon"
                    || attribute_LogicalName == "modifiedonbehalfby")
                {
                    return true;
                }
                return false;
            }
            return false;


        }


        private void sourceDBConnection_txtb_TextChanged(object sender, EventArgs e)
        {
            mySettings.SourceDBConnectionString = sourceDBConnection_txtb.Text;
            SettingsManager.Instance.Save(GetType(), mySettings);
        }
        private void stagingDBConnection_txtb_TextChanged(object sender, EventArgs e)
        {
            mySettings.StagingDBConnectionString = stagingDBConnection_txtb.Text;
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        private void About_btn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This tool is designed to aid migration into D365 using a Dynamics connection and a local MSSQL database.\n\n" +

                            "Creating Staging Tables\n" +
                            "This will create a staging table which is based of the metadata of the selected Dynamics entity. A mapping function enables the Dynamics field types to be mapped to equivalent SQL data types. Each field which contains a guid will have an additional \"_source\" column in the staging database. Additional field are also created to aid with migration.\n\n" +

                            "Steps:\n" +
                            "1. Populate the \"Staging Database Connection String\" with a link to a local MSSQL database, e.g.: Data Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Staging_DMT;Integrated Security=True;Persist Security Info=True;MultipleActiveResultSets=True\n" +
                            "2. Select an entity from the \"Entity\" dropdown.\n" +
                            "3. Click \"Create Staging Table\"\n\n\n" +


                            "Creating Source View Templates\n" +
                            "This will create a source view which is based of the metadata of the selected Dynamics entity. The concept is that the view Template is a starting point, and maps the data to the correct type to work with the other components of this tool. The view can be expanded by adding a Table to call data form, and then by replacing \"NULL\"s in the select with columns from the table. Each line contains a description of the column, either giving the equivalent datatype in dynamics, or if it's a field added to aid with migration, then a description as to how it's used.\n\n" +

                            "Steps:\n" +
                            "1. Populate the \"Source Database Connection String\" with a link to a local MSSQL database, e.g.: Data Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Adventureworks;Integrated Security=True;Persist Security Info=True;MultipleActiveResultSets=True\n" +
                            "2. Select an entity from the \"Entity\" dropdown.\n" +
                            "3. Click \"Create Source View Templates\"");
        }

    }
}