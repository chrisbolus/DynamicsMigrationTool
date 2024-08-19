using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DynamicsMigrationTool
{
    public class SourceToStagingGeneration
    {
        //See Diagram at the bottom for S2S package XML layout.

        private XMLGeneration XMLGen;
        public IOrganizationService Service { get; set; }
        public Settings mySettings { get; set; }
        XNamespace DTS = "www.microsoft.com/SqlServer/Dts";
        string SourceDBId = null;
        string StagingDBId = null;

        //this is the constructor
        public SourceToStagingGeneration(IOrganizationService service, Settings MySettings)
        {
            Service = service;
            mySettings = MySettings;
        }
        internal void UpdateService(IOrganizationService newService)
        {
            Service = newService;
        }
        internal void UpdateSettings(Settings newSettings)
        {
            mySettings = newSettings;
        }



        public void CreatePackage(string entityName)
        {
            var projectName = "SourceToStaging";
            var ssisProjectLocation = mySettings.SourceToStagingLocationString;

            XMLGen = new XMLGeneration(Service, ssisProjectLocation);

            var package = new XDocument();

            var project = XMLGen.GetProjectFile(projectName);

            if (project != null)
            {
                SourceDBId = XMLGen.GetConmgr(ssisProjectLocation, "SourceDB").DTSID;
                StagingDBId = XMLGen.GetConmgr(ssisProjectLocation, "StagingDB").DTSID;

                if (SourceDBId != null && StagingDBId != null)
                {
                    XMLGen.GenerateXML_SSISPackageBase(package, entityName);
                    XMLGen.GenerateXML_Executable_SQLTask_TruncateTable(package, entityName, mySettings.StagingDBSchema, StagingDBId);
                    GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(package, entityName);

                    var dataFlowTaskName = "DataFlowTask";

                    XMLGen.GenerateXML_Executable_DataFlow_Base(package, entityName, dataFlowTaskName);
                    GenerateXML_Executable_DataFlow_SimpleSourceToStaging(package, entityName, dataFlowTaskName);
                    if (mySettings.UseDataTransforms)
                    {
                        GenerateXML_Executable_SQLTask_S2SDataTransforms(package, entityName);
                    }
                    GenerateXML_Executable_SQLTask_S2SReAddPK(package, entityName);
                    GenerateXML_Executable_SQLTask_S2SReAddIndexes(package, entityName);

                    //Constraints link Executables
                    int ConstraintNumber = 1;


                    XMLGen.GenerateXML_Executable_AddConstraint(package, $"Truncate {{{mySettings.StagingDBSchema}}}{{{entityName}}}", "Drop Indexes and PK", ConstraintNumber++);
                    XMLGen.GenerateXML_Executable_AddConstraint(package, "Drop Indexes and PK", dataFlowTaskName, ConstraintNumber++);
                    if(mySettings.UseDataTransforms)
                    {

                        XMLGen.GenerateXML_Executable_AddConstraint(package, dataFlowTaskName, "Data Transforms", ConstraintNumber++);
                        XMLGen.GenerateXML_Executable_AddConstraint(package, "Data Transforms", "ReAdd PK", ConstraintNumber++);
                    }
                    else
                    {
                        XMLGen.GenerateXML_Executable_AddConstraint(package, dataFlowTaskName, "ReAdd PK", ConstraintNumber++);
                    }
                    XMLGen.GenerateXML_Executable_AddConstraint(package, "ReAdd PK", "ReAdd Indexes", ConstraintNumber++);


                    var result = MessageBox.Show($"This will create {entityName}.dtsx at {ssisProjectLocation}\\\n\nWARNING - If that file exists already, it will be OVERWRITTEN!\n\nAre you happy to proceed?", "Warning",
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            XMLGen.SavePackage(package, entityName, ssisProjectLocation);
                            MessageBox.Show("SSIS package created successfully.");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("SSIS package failed to save. Check Source To Staging SSIS Project Location is correct.");
                        }
                        XMLGen.addSSISPackageToProject(entityName, project);
                    }
                }
            }


        }


        private void GenerateXML_Executable_DataFlow_SimpleSourceToStaging(XDocument package, string entityName, string dataFlowTaskName)
        {
            var sourceFields = CRMHelper.GetFullFieldList(Service, entityName);
            var destinationFields = CRMHelper.GetFullFieldList(Service, entityName, true);

            var executable = package.Element(DTS + "Executable").Element(DTS + "Executables").Elements(DTS + "Executable").Where(x => x.Attribute(DTS + "ObjectName").Value == dataFlowTaskName).FirstOrDefault();

            var components = executable.Element(DTS + "ObjectData").Element("pipeline").Element("components");

            XMLGen.GenerateXML_DataFlow_Component_OLEDBSource(components, entityName, mySettings.SourceDBSchema, dataFlowTaskName, sourceFields, "SourceDB");
            XMLGen.GenerateXML_DataFlow_Component_OLEDBDestination(components, entityName, mySettings.StagingDBSchema, dataFlowTaskName, sourceFields, destinationFields, "StagingDB", "OLE DB Source");


            var paths = executable.Element(DTS + "ObjectData").Element("pipeline").Element("paths");

            var pathStringBase = "Package\\" + dataFlowTaskName;

            paths.Add(new XElement("path",
                            new XAttribute("refId", pathStringBase + ".Paths[OLE DB Source Output]"),
                            new XAttribute("name", "OLE DB Source Output"),
                            new XAttribute("startId", pathStringBase + "\\OLE DB Source.Outputs[OLE DB Source Output]"),
                            new XAttribute("endId", pathStringBase + "\\OLE DB Destination.Inputs[OLE DB Destination Input]")
                            ));

        }

        

        private void GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery = $"declare @idxStr nvarchar(2000);\nSELECT @idxStr = (\nselect 'drop index '+s.name+'.'+o.name+'.'+i.name+';'\nfrom sys.indexes i\njoin sys.objects o on i.object_id=o.object_id\njoin sys.schemas as s on s.schema_id = o.schema_id\nwhere o.type <> 'S'\nand i.is_primary_key <> 1\nand i.index_id > 0\nand o.name = '{entityName}'\nand s.name = '{mySettings.StagingDBSchema}'\nFOR xml path('') );\nexec sp_executesql @idxStr;\n\ndeclare @pKStr nvarchar(500);\nSELECT @pKStr = (\nselect 'alter table '+s.name+'.'+o.name+' drop constraint '+i.name+';'\nfrom sys.indexes i\njoin sys.objects o on i.object_id=o.object_id\njoin sys.schemas as s on s.schema_id = o.schema_id\nwhere o.type <> 'S'\nand i.is_primary_key = 1\nand o.name = '{entityName}'\nand s.name = '{mySettings.StagingDBSchema}'\nFOR xml path('') );\nexec sp_executesql @pKStr;\n";

            XMLGen.GenerateXML_Executable_SQLTask_AddTask(package, "Drop Indexes and PK", SQLConnection, SQLQuery);
        }

        private void GenerateXML_Executable_SQLTask_S2SDataTransforms(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery = $"dmt.SP_RunDataTransforms @schemaname = '{mySettings.StagingDBSchema}', @entityname = '{entityName}'";

            XMLGen.GenerateXML_Executable_SQLTask_AddTask(package, "Data Transforms", SQLConnection, SQLQuery);
        }
        private void GenerateXML_Executable_SQLTask_S2SReAddPK(XDocument package, string entityName)
        {
            var PrimaryIdAttribute = Service.GetEntityMetadata(entityName).PrimaryIdAttribute;
            var dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");

            var SQLConnection = StagingDBId;
            var SQLQuery = $"ALTER TABLE [{mySettings.StagingDBSchema}].[{entityName}] ADD CONSTRAINT [PK_{entityName}_{dateTimeNow}] PRIMARY KEY CLUSTERED \n(\n\t[Source_System_Id] ASC,\n\t[{PrimaryIdAttribute}_Source] ASC\n) ON [PRIMARY]\nGO";

            XMLGen.GenerateXML_Executable_SQLTask_AddTask(package, "ReAdd PK", SQLConnection, SQLQuery);
        }
        private void GenerateXML_Executable_SQLTask_S2SReAddIndexes(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery =
                                //$"CREATE NONCLUSTERED INDEX [IDX_CreateParameters] ON [{mySettings.StagingDBSchema}].[{entityName}]\n(\n\t[FlagCreate] ASC,\n\t[DynCreateId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_UpdateParameters] ON [{mySettings.StagingDBSchema}].[{entityName}]\n(\n\t[FlagUpdate] ASC,\n\t[DynUpdateId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_DeleteParameters] ON [{mySettings.StagingDBSchema}].[{entityName}]\n(\n\t[FlagDelete] ASC,\n\t[DynDeleteId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_Processing_Status] ON [{mySettings.StagingDBSchema}].[{entityName}]\n(\n\t[Processing_Status] ASC,\n\t[FlagCreate] ASC,\n\t[DynCreateId] ASC,\n\t[FlagUpdate] ASC,\n\t[DynUpdateId] ASC,\n\t[FlagDelete] ASC,\n\t[DynDeleteId] ASC\n) ON [PRIMARY]\nGO";
                                $"CREATE NONCLUSTERED INDEX [IDX_Processing_Status] ON [{mySettings.StagingDBSchema}].[{entityName}]\n(\n\t[Processing_Status] ASC,\n\t[DynId] ASC\n) ON [PRIMARY]\nGO";

            XMLGen.GenerateXML_Executable_SQLTask_AddTask(package, "ReAdd Indexes", SQLConnection, SQLQuery);
        }


    }
}

//XML S2S PACKAGE LAYOUT

//<DTS:Executable 
//  <DTS:Property
//  <DTS:Variables />
//  <DTS:Executables>
//    <DTS:Executable
//      <DTS:Variables />
//      <DTS:ObjectData>
//        <pipeline
//          <components>
//
//            ////Source
//            <component
//              <properties></properties>
//              <connections></connections>
//              <outputs>
//                <output
//                  <outputColumns></outputColumns>
//                  <externalMetadataColumns></externalMetadataColumns>
//                </output>
//                <output
//                  <outputColumns></outputColumns>
//                </output>
//              </outputs>
//            </component>
//
//            ////Staging
//            <component
//              <properties> </properties>
//              <connections> </connections>
//              <inputs>
//                <input
//                  <inputColumns></inputColumns>
//                  <externalMetadataColumns></externalMetadataColumns>
//                </input>
//              </inputs>
//              <outputs>
//                <output
//                  <outputColumns></outputColumns>
//                </output>
//              </outputs>
//            </component>
//
//          </components>
//          <paths></paths>
//        </pipeline>
//      </DTS:ObjectData>
//    </DTS:Executable>
//  </DTS:Executables>
//</ DTS:Executable >