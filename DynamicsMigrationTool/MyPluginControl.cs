﻿using McTools.Xrm.Connection;
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
using DynamicsMigrationTool;
using ScintillaNET;

namespace DynamicsMigrationTool
{
    public partial class MyPluginControl : PluginControlBase
    {
        private Settings mySettings;
        private SourceToStagingGeneration sourceToStagingGeneration;

        public MyPluginControl()
        {
            InitializeComponent();
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
                    sourceToStagingGeneration = new SourceToStagingGeneration(Service, mySettings);
                }

                sourceDBConnection_txtb.Text = mySettings.SourceDBConnectionString;
                stagingDBConnection_txtb.Text = mySettings.StagingDBConnectionString;
                sourceToStagingLocation_txtb.Text = mySettings.SourceToStagingLocationString;

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


            if(sourceToStagingGeneration != null)
            {
                sourceToStagingGeneration.UpdateService(newService);
            }
            else 
            {
                sourceToStagingGeneration = new SourceToStagingGeneration(newService, mySettings);
            }
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

                var sourceDBConnectionString = new SqlConnection();

                Boolean isSourceDBConnectionValid = true;
                try
                {
                    sourceDBConnectionString = new SqlConnection(mySettings.SourceDBConnectionString);
                    sourceDBConnectionString.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Please review Source Database Connection String:\n{mySettings.SourceDBConnectionString}\n\nError: {ex.Message}\n\nExample Connection String:\nData Source=DESKTOP\\SQLEXPRESS;Initial Catalog=Source_DB;Integrated Security=True;");
                    isSourceDBConnectionValid = false;
                }

                if (isSourceDBConnectionValid)
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
            var fieldList = CRMHelper.GetFullFieldList(Service, entity);
            var isFirstColumn = true;

            using (var command = connection.CreateCommand())
            {
                command.CommandTimeout = 0;
                command.CommandText = $@"
CREATE OR ALTER VIEW [dmt].{entity.LogicalName}_Template AS 
--RENAME THIS VIEW! Remove the ""_Template\"" suffix, or replace it with your own suffix.
--THIS IS A BASE VIEW OF ALL FIELDS FOR THE {entity.LogicalName.ToUpper()} ENTITY WITH ALL THE CASTS TO ENSURE ALL THE DATA ENDS UP IN THE CORRECT FORMAT. ADD A FROM STATEMENT TO PULL DATA FROM THE TABLE YOU WANT TO USE, AND REPLACE THE NULLS IN THE CASTS WITH THE FIELDS FROM YOUR SOURCE TABLE .

SELECT
";


                foreach (var field in fieldList)
                {
                    var defaultValue = "NULL";
                    var targetListMessage = field.DynEntityLookupTargets_List == null ? "" : $" Target entity(s) = {field.DynEntityLookupTargets_List}.";
                    var commentText = $"Dynamics Attribute Type = {field.DynDataType_Readable}.{targetListMessage}";

                    if (field.isDMTField == true)
                    {
                        if (field.isLookup)
                        {
                            if (field.isPrimaryKey)
                            {
                                if (field.fieldName == field.FieldDerivedFrom + "_Source")
                                {
                                    commentText = $"DMT Field, used to handle the id of the record from the source system. Staging table requires this to be populated and unique when combined with Source_System_Id.";
                                }
                                else
                                {
                                    commentText = $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with Leader_System_Id.";
                                }
                            }
                            else
                            {
                                commentText = $"DMT Field, used to handle the source system ids of other entities.{targetListMessage}";
                            }
                        }

                        else if (field.fieldName == "Source_System_Id")
                        {
                            defaultValue = "1";
                            commentText = "DMT Field, used to handle multiple source systems. Default value = 1. Change if using a second (etc.) source system.";
                        }
                        else if (field.fieldName == "Processing_Status")
                        {
                            defaultValue = "1";
                            commentText = "DMT Field, used to handle relationship mapping. Default value = 1. Set to 0 for records that shouldn't be loaded.";
                        }
                        else if (field.fieldName == "Leader_System_Id")
                        {
                            commentText = $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with {entity.PrimaryIdAttribute}_Leader.";
                        }
                    }

                    command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST({defaultValue} AS {field.DBDataType}) AS {field.fieldName}", commentText, isFirstColumn);

                    isFirstColumn = false;
                }

                command.CommandText += "\n--FROM ";

                command.ExecuteNonQuery();
            }

        }

        public void CreateTemplateSourceView_Old(SqlConnection connection, EntityMetadata entity)
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

                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(1 AS INT) AS Source_System_Id",
                                                                        $"DMT Field, used to handle multiple source systems. Default value = 1. Change if using a second (etc.) source system.", true);
                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(1 AS INT) AS Processing_Status",
                                                                        $"DMT Field, used to handle relationship mapping. Default value = 1. Set to 0 for records that shouldn't be loaded.");


                foreach (var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
                {
                    EntityAttribute_AdditionalInfo entAAI = CRMHelper.Get_EntityAttribute_AdditionalInfo(Service, attribute);


                    if (entAAI.isValidForMigration)
                    {
                        var targetListMessage = entAAI.DynEntityLookupTargets_List == null ? "" : $" Target entity(s) = {entAAI.DynEntityLookupTargets_List}.";

                        command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(NULL AS {entAAI.DBDataType}) AS {attribute.LogicalName}",
                                                                                $"Dynamics Attribute Type = {entAAI.DynDataType_Readable}.{targetListMessage}");
                        if (entAAI.isLookup)
                        {
                            if (entAAI.isPrimaryKey)
                            {
                                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Source",
                                                                                        $"DMT Field, used to handle the id of the record from the source system. Staging table requires this to be populated and unique when combined with Source_System_Id.");
                                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Leader",
                                                                                        $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with Leader_System_Id.");
                                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(NULL AS INT) AS Leader_System_Id",
                                                                                        $"DMT Field, used to handle relationship mapping to another {entity.LogicalName} record. For Staging view dbo.{entity.LogicalName}_RelationshipMap to work, this field must be used in conjuction with {attribute.LogicalName}_Leader.");
                            }
                            else
                            {
                                command.CommandText += TemplateSourceView_HandleCommentSpacing($"CAST(NULL AS NVARCHAR(255)) AS {attribute.LogicalName}_Source",
                                                                                        $"DMT Field, used to handle the source system ids of other entities.{targetListMessage}");
                            }
                        }
                    }
                }

                command.CommandText += "\n--FROM ";

                command.ExecuteNonQuery();
            }

        }

        public string TemplateSourceView_HandleCommentSpacing(string queryLine, string comment, bool isFirstColumn = false)
        {
            int spacing = 80;
            int spacingoffset = queryLine.Length >= spacing ? 10 : spacing - queryLine.Length;
            string addComma = isFirstColumn ? " " : ",";

            return $"       {addComma}{queryLine}{new String(' ', spacingoffset)}--{comment}\n";
        }



        public void CreateStagingTable(SqlConnection connection, EntityMetadata entity)
        {
            var fieldList = CRMHelper.GetFullFieldList(Service, entity, true);

            using (var command = connection.CreateCommand())
            {
                command.CommandTimeout = 0;
                command.CommandText = $@"
                    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].{entity.LogicalName}') AND type in (N'U'))
                    DROP TABLE [dbo].{entity.LogicalName}

                    CREATE TABLE [dbo].{entity.LogicalName} (";

                foreach (var field in fieldList)
                {
                    command.CommandText = command.CommandText + $"{field.fieldName} {field.DBDataType} {(field.stgNotNull ? " NOT NULL" : "")},\n";
                }

                command.CommandText = command.CommandText + "\n)";

                command.ExecuteNonQuery();
            }
        

            CreateIndexes(connection, entity);
            CreateRelationshipMapView(connection, entity);
        }

        private void CreateIndexes(SqlConnection connection, EntityMetadata entity)
        {
            var dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");

            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT PK_{entity.LogicalName}_{dateTimeNow} PRIMARY KEY CLUSTERED (Source_System_Id ASC, {entity.PrimaryIdAttribute}_Source ASC) ON [PRIMARY]
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT [CN_{entity.LogicalName}_FlagCreate_{dateTimeNow}]  DEFAULT (1) FOR [FlagCreate]
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT [CN_{entity.LogicalName}_FlagUpdate_{dateTimeNow}]  DEFAULT (0) FOR [FlagUpdate]
                ALTER TABLE dbo.{entity.LogicalName} ADD CONSTRAINT [CN_{entity.LogicalName}_FlagDelete_{dateTimeNow}]  DEFAULT (0) FOR [FlagDelete]
                CREATE INDEX IDX_Processing_Status ON dbo.{entity.LogicalName}(Processing_Status, FlagCreate, DynCreateId, FlagUpdate, DynUpdateId, FlagDelete, DynDeleteId);
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
        private void sourceToStagingLocation_txtb_TextChanged(object sender, EventArgs e)
        {
            mySettings.SourceToStagingLocationString = sourceToStagingLocation_txtb.Text;
            SettingsManager.Instance.Save(GetType(), mySettings);

            if (sourceToStagingGeneration != null)
            {
                sourceToStagingGeneration.UpdateSettings(mySettings);
            }
            else
            {
                sourceToStagingGeneration = new SourceToStagingGeneration(Service, mySettings);
            }
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

        public void test_btn_Click(object sender, EventArgs e)
        {
        }

        private void CreateS2SPackage_Btn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(TestConnection);

            if (IsEntitySelected())
            {
                var entityMetadata_noattr = (EntityMetadata)EntityCmb.SelectedItem;

                sourceToStagingGeneration.CreatePackage(entityMetadata_noattr.LogicalName);
            }

        }
    }
}