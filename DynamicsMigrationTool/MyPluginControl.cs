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

namespace DynamicsMigrationTool
{
    public partial class MyPluginControl : PluginControlBase
    {
        private Settings mySettings;

        public MyPluginControl()
        {
            InitializeComponent();
        }

        private class Attribute_StagingInfo
        {
            public string DBDataType;
            public bool CreateColumn = true;
            public bool CreateSourceField = false;
            public string DynEntityLookupTargets_List;
            public string DynDataType_Readable;
        }


        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            ShowInfoNotification("This is a notification that can lead to XrmToolBox repository", new Uri("https://github.com/MscrmTools/XrmToolBox"));

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

                Boolean isStagingDBConnectionValid = true;
                try
                {
                    stagingDBConnectionString = new SqlConnection(mySettings.StagingDBConnectionString);
                    stagingDBConnectionString.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Please review Staging Database Connection String:\n{mySettings.StagingDBConnectionString}\n\nError: {ex.Message}\n\nExample Connection String:\nData Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Staging_DB;Integrated Security=True;");
                    isStagingDBConnectionValid = false;
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

                var stagingDBConnectionString = new SqlConnection();

                Boolean isStagingDBConnectionValid = true;
                try
                {
                    stagingDBConnectionString = new SqlConnection(mySettings.StagingDBConnectionString);
                    stagingDBConnectionString.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Please review Staging Database Connection String:\n{mySettings.StagingDBConnectionString}\n\nError: {ex.Message}\n\nExample Connection String:\nData Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Staging_DB;Integrated Security=True;");
                    isStagingDBConnectionValid = false;
                }

                if (isStagingDBConnectionValid)
                {
                    CreateTemplateSourceView(stagingDBConnectionString, entityMetadata);

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
                    if(attribute.IsValidForCreate == true || attribute.IsValidForUpdate == true)
                    {
                        Attribute_StagingInfo attSI = Get_Attribute_StagingInfo(attribute);

                        var targetListMessage = attSI.DynEntityLookupTargets_List == null ? "" : $" Target entity(s) = { attSI.DynEntityLookupTargets_List}.";

                        if (attSI.CreateColumn && (attribute.AttributeOf == null || (attribute.IsLogical == null || !attribute.IsLogical.Value)))
                        {
                            command.CommandText += TemplateSourceView_HandleSpacing($"CAST(NULL AS {attSI.DBDataType}) AS {attribute.LogicalName}", 
                                                                                    $"Dynamics Attribute Type = {attSI.DynDataType_Readable}.{targetListMessage}");
                        }
                        if (attSI.CreateSourceField)
                        {
                            if (attribute.LogicalName == entity.PrimaryIdAttribute)
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
                //removing final comma
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
                    FlagCreate BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagCreate DEFAULT 1, 
                    FlagUpdate BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagUpdate DEFAULT 0, 
                    FlagDelete BIT NOT NULL CONSTRAINT CN_{entity.LogicalName}_FlagDelete DEFAULT 0, 
                    DateCreate DATETIME, 
                    DateUpdate DATETIME, 
                    DateDelete DATETIME, 
                    DynCreateId UNIQUEIDENTIFIER, 
                    DynUpdateId UNIQUEIDENTIFIER, 
                    DynDeleteId UNIQUEIDENTIFIER 
                    );" + "\n\n";



                foreach (var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
                {
                    if (attribute.IsValidForCreate == true || attribute.IsValidForUpdate == true)
                    {
                        Attribute_StagingInfo attSI = Get_Attribute_StagingInfo(attribute);

                        if (attSI.CreateColumn && (attribute.AttributeOf == null || (attribute.IsLogical == null || !attribute.IsLogical.Value)))
                        {
                            command.CommandText += $"ALTER TABLE dbo.{entity.LogicalName} ADD {attribute.LogicalName} {attSI.DBDataType};\n";
                        }
                        if (attSI.CreateSourceField)
                        {
                            if (attribute.LogicalName == entity.PrimaryIdAttribute)
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

        private Attribute_StagingInfo Get_Attribute_StagingInfo(AttributeMetadata attribute)
        {
            var Stg = new Attribute_StagingInfo();
            Stg.DynDataType_Readable = attribute.AttributeType?.ToString();      //this is modified in a few of cases below.

            if (attribute.LogicalName == "EntityImage")
            {
                Stg.DBDataType = "IMAGE";
                Stg.DynDataType_Readable = attribute.LogicalName;
            }
            if (attribute.AttributeType == AttributeTypeCode.BigInt)
            {
                Stg.DBDataType = "BIGINT";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Boolean 
                || attribute.AttributeType == AttributeTypeCode.EntityName 
                || attribute.AttributeType == AttributeTypeCode.Picklist 
                || attribute.AttributeType == AttributeTypeCode.State 
                || attribute.AttributeType == AttributeTypeCode.Status)
            {
                Stg.DBDataType = "NVARCHAR(255)";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Customer
                || attribute.AttributeType == AttributeTypeCode.Lookup
                || attribute.AttributeType == AttributeTypeCode.Owner
                || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier)
            {
                Stg.DBDataType = "UNIQUEIDENTIFIER";
                Stg.CreateSourceField = true;
                Stg.DynEntityLookupTargets_List = GetEntityLookupTargetsList(attribute);
            }
            else if (attribute.AttributeType == AttributeTypeCode.DateTime)
            {
                Stg.DBDataType = "DATETIME";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Decimal)
            {
                Stg.DBDataType = "DECIMAL(23, 10)";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Double)
            {
                Stg.DBDataType = "FLOAT";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Integer)
            {
                Stg.DBDataType = "INT";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Memo)
            {
                Stg.DBDataType = "NVARCHAR(MAX)";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Money)
            {
                Stg.DBDataType = "MONEY";
            }
            else if (attribute.AttributeType == AttributeTypeCode.String)
            {
                var strlength = (int)attribute.GetType().GetProperty("MaxLength").GetValue(attribute);
                Stg.DynDataType_Readable = Stg.DynDataType_Readable + $"({strlength})";
                if (strlength > 4000)
                {
                    Stg.DBDataType = "NVARCHAR(MAX)";
                }
                else
                {
                    Stg.DBDataType = $"NVARCHAR({strlength})";
                }
            }
            else if (attribute.AttributeType == AttributeTypeCode.Virtual)
            {
                if (attribute.AttributeTypeName.Value == "MultiSelectPicklistType")
                {
                    Stg.DBDataType = "NVARCHAR(1000)";
                    Stg.DynDataType_Readable = Stg.DynDataType_Readable + " (MultiSelectPicklistType)";
                }
                else
                {
                    Stg.CreateColumn = false;
                }
            }
            else
            {
                Stg.CreateColumn = false;
                Stg.DBDataType = null;
            }

            return Stg;
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
            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT PK_{entity.LogicalName} PRIMARY KEY CLUSTERED (Source_System_Id ASC, {entity.PrimaryIdAttribute}_Source ASC) ON [PRIMARY]
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
                            "3. Click \"Create Staging Table\"");
        }

    }
}