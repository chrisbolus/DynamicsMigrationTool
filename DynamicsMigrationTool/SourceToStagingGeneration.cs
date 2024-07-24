using Microsoft.SqlServer.Management.Smo;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Metadata;
using ScintillaNET;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Shapes;
using System.Xml.Linq;
using static Azure.Core.HttpHeader;
using static ScintillaNET.Style;

namespace DynamicsMigrationTool
{
    public class SourceToStagingGeneration
    {
        //See Diagram at the bottom for S2S package XML layout.

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
            var package = new XDocument();
            var packageLocation = mySettings.SourceToStagingLocationString;

            var project = GetS2SProject();

            if (project != null)
            {
                SourceDBId = GetConmgrId(packageLocation, "SourceDB");
                StagingDBId = GetConmgrId(packageLocation, "StagingDB");

                if (SourceDBId != null && StagingDBId != null)
                {
                    GenerateXML_SSISPackageBase(package, entityName);
                    GenerateXML_Executable_SQLTask_S2STruncateStagingTable(package, entityName);
                    GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(package, entityName);
                    GenerateXML_Executable_DataFlow_SimpleS2S(package, entityName);
                    GenerateXML_Executable_SQLTask_S2SReAddPK(package, entityName);
                    GenerateXML_Executable_SQLTask_S2SReAddIndexes(package, entityName);

                    //Constraints link Executables
                    int ConstraintNumber = 1;
                    GenerateXML_Executable_SQLTask_AddConstraint(package, "Truncate Staging Table", "Drop Indexes and PK", ConstraintNumber++);
                    GenerateXML_Executable_SQLTask_AddConstraint(package, "Drop Indexes and PK", "Data Flow Task_1", ConstraintNumber++);
                    GenerateXML_Executable_SQLTask_AddConstraint(package, "Data Flow Task_1", "ReAdd PK", ConstraintNumber++);
                    GenerateXML_Executable_SQLTask_AddConstraint(package, "ReAdd PK", "ReAdd Indexes", ConstraintNumber++);

                    var result = MessageBox.Show($"This will create {entityName}.dtsx at {packageLocation}\\\n\nWARNING - If that file exists already, it will be OVERWRITTEN!\n\nAre you happy to proceed?", "Warning",
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            SavePackage(package, entityName, packageLocation);
                            MessageBox.Show("SSIS package created successfully.");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("SSIS package failed to save. Check Source To Staging SSIS Project Location is correct.");
                        }

                        addSSISPackageToProject(entityName);
                    }
                }
            }

            
        }

        private string GetConmgrId(string packageLocation, string ConnectionName)
        {
            var packagePath = packageLocation + $"\\{ConnectionName}.conmgr";
            var conmgr = new XDocument();
            string conmgrid = null;

            Boolean goodPath = false;

            try
            {
                conmgr = XDocument.Load(packagePath);
                goodPath = true;
            }
            catch (Exception e)
            {
                MessageBox.Show($"Unable to find a file called {ConnectionName}.conmgr at the Source To Staging SSIS Project Location."); 
            }

            if(goodPath)
            {
                try
                {
                    conmgrid = conmgr.Element(DTS + "ConnectionManager").Attribute(DTS + "DTSID").Value;
                }
                catch (Exception e)
                {
                    MessageBox.Show($"{ConnectionName}.conmgr file layout is incorrect, unable to find the DTS:DTSID.");
                }
            }
            
            return conmgrid;            
        }

        private void GenerateXML_SSISPackageBase(XDocument package, string entityName)
        {
            package.Add(new XElement(DTS + "Executable",
                            new XAttribute(XNamespace.Xmlns + "DTS", DTS.ToString()),
                            new XAttribute(DTS + "ObjectName", entityName),
                            new XAttribute(DTS + "refId", "Package"),
                            new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                            new XAttribute(DTS + "ExecutableType", "Microsoft.Package"),
                            new XAttribute(DTS + "CreationName", "Microsoft.Package"),
                            new XAttribute(DTS + "VersionBuild", "1"),
                            new XAttribute(DTS + "VersionGUID", GenerateNewXMLGuid()),
                            new XElement(DTS + "Property", "8",
                                new XAttribute(DTS + "Name", "PackageFormatVersion")
                            ),
                            new XElement(DTS + "Variables"),
                            new XElement(DTS + "Executables"),
                            new XElement(DTS + "PrecedenceConstraints")
                            ));
        }

        private void GenerateXML_Executable_SQLTask_S2STruncateStagingTable(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery = $"truncate table dbo.{entityName}";

            GenerateXML_Executable_SQLTask_AddTask(package, "Truncate Staging Table", SQLConnection, SQLQuery);
        }

        private void GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery = $"declare @idxStr nvarchar(2000);\nSELECT @idxStr = (\nselect 'drop index '+o.name+'.'+i.name+';'\nfrom sys.indexes i\njoin sys.objects o on i.object_id=o.object_id\njoin sys.schemas as s on s.schema_id = o.schema_id\nwhere o.type <> 'S'\nand i.is_primary_key <> 1\nand i.index_id > 0\nand o.name = '{entityName}'\nand s.name = 'dbo'\nFOR xml path('') );\nexec sp_executesql @idxStr;\n\ndeclare @pKStr nvarchar(500);\nSELECT @pKStr = (\nselect 'alter table '+o.name+' drop constraint '+i.name+';'\nfrom sys.indexes i\njoin sys.objects o on i.object_id=o.object_id\njoin sys.schemas as s on s.schema_id = o.schema_id\nwhere o.type <> 'S'\nand i.is_primary_key = 1\nand o.name = '{entityName}'\nand s.name = 'dbo'\nFOR xml path('') );\nexec sp_executesql @pKStr;\n";

            GenerateXML_Executable_SQLTask_AddTask(package, "Drop Indexes and PK", SQLConnection, SQLQuery);
        }
        private void GenerateXML_Executable_SQLTask_S2SReAddPK(XDocument package, string entityName)
        {
            var PrimaryIdAttribute = Service.GetEntityMetadata(entityName).PrimaryIdAttribute;
            var dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");

            var SQLConnection = StagingDBId;
            var SQLQuery = $"ALTER TABLE [dbo].[{entityName}] ADD CONSTRAINT [PK_{entityName}_{dateTimeNow}] PRIMARY KEY CLUSTERED \n(\n\t[Source_System_Id] ASC,\n\t[{PrimaryIdAttribute}_Source] ASC\n) ON [PRIMARY]\nGO";

            GenerateXML_Executable_SQLTask_AddTask(package, "ReAdd PK", SQLConnection, SQLQuery);
        }
        private void GenerateXML_Executable_SQLTask_S2SReAddIndexes(XDocument package, string entityName)
        {
            var SQLConnection = StagingDBId;
            var SQLQuery =
                                //$"CREATE NONCLUSTERED INDEX [IDX_CreateParameters] ON [dbo].[{entityName}]\n(\n\t[FlagCreate] ASC,\n\t[DynCreateId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_UpdateParameters] ON [dbo].[{entityName}]\n(\n\t[FlagUpdate] ASC,\n\t[DynUpdateId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_DeleteParameters] ON [dbo].[{entityName}]\n(\n\t[FlagDelete] ASC,\n\t[DynDeleteId] ASC\n) ON [PRIMARY]\nGO\n\n" +
                                //$"CREATE NONCLUSTERED INDEX [IDX_Processing_Status] ON [dbo].[{entityName}]\n(\n\t[Processing_Status] ASC,\n\t[FlagCreate] ASC,\n\t[DynCreateId] ASC,\n\t[FlagUpdate] ASC,\n\t[DynUpdateId] ASC,\n\t[FlagDelete] ASC,\n\t[DynDeleteId] ASC\n) ON [PRIMARY]\nGO";
                                $"CREATE NONCLUSTERED INDEX [IDX_Processing_Status] ON [dbo].[{entityName}]\n(\n\t[Processing_Status] ASC,\n\t[DynId] ASC\n) ON [PRIMARY]\nGO";

            GenerateXML_Executable_SQLTask_AddTask(package, "ReAdd Indexes", SQLConnection, SQLQuery);
        }

        private void GenerateXML_Executable_SQLTask_AddTask(XDocument package, string executableName, string SQLConnection, string SQLQuery)
        {
            XNamespace sqlTask = "www.microsoft.com/sqlserver/dts/tasks/sqltask";

            var executables = package.Element(DTS + "Executable").Element(DTS + "Executables");

            XElement executable = new XElement(DTS + "Executable",
                new XAttribute(DTS + "refId", $"Package\\{executableName}"),
                new XAttribute(DTS + "ObjectName", executableName),
                new XAttribute(DTS + "CreationName", "Microsoft.ExecuteSQLTask"),
                new XAttribute(DTS + "Description", "Execute SQL Task"),
                new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                new XAttribute(DTS + "ExecutableType", "Microsoft.ExecuteSQLTask"),
                new XAttribute(DTS + "LocaleID", "-1"),
                new XAttribute(DTS + "ThreadHint", "0"),
                new XElement(DTS + "Variables"),
                new XElement(DTS + "ObjectData",
                    new XElement(sqlTask + "SqlTaskData",
                        new XAttribute(XNamespace.Xmlns + "sqlTask", sqlTask.ToString()),
                        new XAttribute(sqlTask + "Connection", SQLConnection),
                        new XAttribute(sqlTask + "SqlStatementSource", SQLQuery)
                    )
                )
            );

            executables.Add(executable);
        }

        private void GenerateXML_Executable_SQLTask_AddConstraint(XDocument package, string executableName_from, string executableName_to, int ConstraintNumber)
        {
            var precedenceConstraints = package.Element(DTS + "Executable").Element(DTS + "PrecedenceConstraints");

            XElement precedenceConstraint = new XElement(DTS + "PrecedenceConstraint",
                new XAttribute(DTS + "From", $"Package\\{executableName_from}"),
                new XAttribute(DTS + "To",   $"Package\\{executableName_to}"),
                new XAttribute(DTS + "refId", $"Package.PrecedenceConstraints[Constraint_{ConstraintNumber}]"),
                new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                new XAttribute(DTS + "LogicalAnd", "True"),
                new XAttribute(DTS + "ObjectName", "Constraint"),
                new XAttribute(DTS + "CreationName", string.Empty)
            );

            precedenceConstraints.Add(precedenceConstraint);
        }

        private void GenerateXML_Executable_DataFlow_SimpleS2S(XDocument package, string entityName)
        {
            int dataFlowNumber = 1;

            var executables = package.Element(DTS + "Executable").Element(DTS + "Executables");

            var executable = new XElement(DTS + "Executable",
                                new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                                new XAttribute(DTS + "refId", "Package\\Data Flow Task_" + dataFlowNumber),
                                new XAttribute(DTS + "CreationName", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "ExecutableType", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "ObjectName", "Data Flow Task_" + dataFlowNumber),
                                new XElement(DTS + "ObjectData",
                                    new XElement("pipeline",
                                        new XAttribute("version", "1"),
                                        new XElement("components"),
                                        new XElement("paths")
                                        )));

            executables.Add(executable);

            var components = executable.Element(DTS + "ObjectData").Element("pipeline").Element("components");

            GenerateXML_DataFlow_Component_Source(components, entityName, dataFlowNumber);
            GenerateXML_DataFlow_Component_Staging(components, entityName, dataFlowNumber);

            var paths = executable.Element(DTS + "ObjectData").Element("pipeline").Element("paths");

            var pathStringBase = "Package\\Data Flow Task_" + dataFlowNumber;

            paths.Add(new XElement("path",
                            new XAttribute("refId", pathStringBase + ".Paths[OLE DB Source Output]"),
                            new XAttribute("name", "OLE DB Source Output"),
                            new XAttribute("startId", pathStringBase + "\\OLE DB Source.Outputs[OLE DB Source Output]"),
                            new XAttribute("endId", pathStringBase + "\\OLE DB Destination.Inputs[OLE DB Destination Input]")
                            ));
        }

        private void GenerateXML_DataFlow_Component_Source(XElement components, string entityName, int dataFlowNumber)
        {
            var component = new XElement("component",
                                new XAttribute("refId", $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Source"),
                                new XAttribute("componentClassID", "Microsoft.OLEDBSource"),
                                new XAttribute("name", "OLE DB Source"),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "7")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_Properties(component, entityName);
            GenerateXML_DataFlow_Component_Connections(component, entityName, dataFlowNumber);
            GenerateXML_DataFlow_Component_Source_Output(component, entityName, dataFlowNumber);
        }


        private void GenerateXML_DataFlow_Component_Staging(XElement components, string entityName, int dataFlowNumber)
        {
            var component = new XElement("component",
                                new XAttribute("refId", $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Destination"),
                                new XAttribute("componentClassID", "Microsoft.OLEDBDestination"),
                                new XAttribute("name", "OLE DB Destination"),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "4")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_Properties(component, entityName);
            GenerateXML_DataFlow_Component_Connections(component, entityName, dataFlowNumber);
            GenerateXML_DataFlow_Component_Staging_Input(component, entityName, dataFlowNumber);
            GenerateXML_DataFlow_Component_Staging_Output(component, entityName, dataFlowNumber);
        }



        private void GenerateXML_DataFlow_Component_Source_Output(XElement component, string entityName, int dataFlowNumber)
        {
            var outputs = new XElement("outputs");
            component.Add(outputs);

            var data_path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Source.Outputs[OLE DB Source Output]";
            var error_path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Source.Outputs[OLE DB Source Error Output]";


            var data_Output = new XElement("output",
                                    new XAttribute("refId", data_path),
                                    new XAttribute("name", "OLE DB Source Output"));

            var error_Output = new XElement("output",
                                    new XAttribute("refId", error_path),
                                    new XAttribute("isErrorOut", "true"),
                                    new XAttribute("name", "OLE DB Source Error Output"));

            outputs.Add(data_Output);           
            outputs.Add(error_Output);           

            var data_OutputColumns = new XElement("outputColumns");
            var data_ExternalMetadataColumns = new XElement("externalMetadataColumns",
                                                    new XAttribute("isUsed", "True"));
            var error_OutputColumns = new XElement("outputColumns");

            GenerateXML_DataFlow_Component_AddErrorOutputColumns(error_OutputColumns, error_path);




            var entityMetadata = Service.GetEntityMetadata(entityName);

            var EntAAIList = CRMHelper.GetFullFieldList(Service, entityMetadata);

            foreach (var EntAAI in EntAAIList)
            {
                GenerateXML_DataFlow_Component_OutputColumn(data_OutputColumns, EntAAI, data_path);
                GenerateXML_DataFlow_Component_externalMetadataColumn(data_ExternalMetadataColumns, EntAAI, data_path);

                GenerateXML_DataFlow_Component_OutputColumn(error_OutputColumns, EntAAI, error_path, true);
            }

            data_Output.Add(data_OutputColumns);
            data_Output.Add(data_ExternalMetadataColumns);
            error_Output.Add(error_OutputColumns);

        }


        private void GenerateXML_DataFlow_Component_InputColumn(XElement data_InputColumns, EntityAttribute_AdditionalInfo entAAI, string path, string lineage_path)
        {
            XElement inputColumn =
                new XElement("inputColumn",

                    new XAttribute("cachedDataType", entAAI.SSISDataType),
                    new XAttribute("cachedName", entAAI.fieldName),

                    new XAttribute("refId", $"{path}.Columns[{entAAI.fieldName}]"),
                    new XAttribute("lineageId", $"{lineage_path}.Columns[{entAAI.fieldName}]"),
                    new XAttribute("externalMetadataColumnId", $"{path}.ExternalColumns[{entAAI.fieldName}]")
                );
            
            if (entAAI.SSISDataType == "wstr")
            {
                inputColumn.Add(new XAttribute("cachedLength", entAAI.StringLength == -1 ? "" : entAAI.StringLength.ToString()));
            }
            if (entAAI.SSISDataType == "numeric")
            {
                inputColumn.Add(new XAttribute("cachedPrecision", "23"));
                inputColumn.Add(new XAttribute("cachedScale", "10"));
            }
            data_InputColumns.Add(inputColumn);
        }

        private void GenerateXML_DataFlow_Component_OutputColumn(XElement data_OutputColumns, EntityAttribute_AdditionalInfo entAAI, string path, bool isErrorOutput = false)
        {
            XElement outputColumn =
                new XElement("outputColumn",

                    new XAttribute("dataType", entAAI.SSISDataType),
                    new XAttribute("name", entAAI.fieldName),

                    new XAttribute("refId", $"{path}.Columns[{entAAI.fieldName}]"),
                    new XAttribute("lineageId", $"{path}.Columns[{entAAI.fieldName}]")
                );
            if (isErrorOutput == false)
            {
                outputColumn.Add(new XAttribute("externalMetadataColumnId", $"{path}.ExternalColumns[{entAAI.fieldName}]"));

                outputColumn.Add(new XAttribute("errorOrTruncationOperation", "Conversion"));
                outputColumn.Add(new XAttribute("errorRowDisposition", "FailComponent"));
                outputColumn.Add(new XAttribute("truncationRowDisposition", "FailComponent"));
            }
            if (entAAI.SSISDataType == "wstr")
            {
                outputColumn.Add(new XAttribute("length", entAAI.StringLength == -1 ? "" : entAAI.StringLength.ToString()));
            }
            if (entAAI.SSISDataType == "numeric")
            {
                outputColumn.Add(new XAttribute("precision", "23"));
                outputColumn.Add(new XAttribute("scale", "10"));
            }
            data_OutputColumns.Add(outputColumn);
        }
        private void GenerateXML_DataFlow_Component_externalMetadataColumn(XElement data_ExternalMetadataColumns, EntityAttribute_AdditionalInfo entAAI, string path)
        {
            XElement externalMetadataColumn =
                new XElement("externalMetadataColumn",

                    new XAttribute("dataType", entAAI.SSISDataType),
                    new XAttribute("name", entAAI.fieldName),
                    new XAttribute("refId", $"{path}.ExternalColumns[{entAAI.fieldName}]")
                );
            if (entAAI.SSISDataType == "wstr")
            {
                externalMetadataColumn.Add(new XAttribute("length", entAAI.StringLength == -1 ? "" : entAAI.StringLength.ToString()));
            }
            if (entAAI.SSISDataType == "numeric")
            {
                externalMetadataColumn.Add(new XAttribute("cachedPrecision", "23"));
                externalMetadataColumn.Add(new XAttribute("cachedScale", "10"));
            }
            data_ExternalMetadataColumns.Add(externalMetadataColumn);
        }

        private void GenerateXML_DataFlow_Component_Staging_Input(XElement component, string entityName, int dataFlowNumber)
        {
            var inputs = new XElement("inputs");
            component.Add(inputs);

            var data_path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Destination.Inputs[OLE DB Destination Input]";
            var lineage_path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Source.Outputs[OLE DB Source Output]";




            var data_Input = new XElement("input",
                                    new XAttribute("refId", data_path),
                                    new XAttribute("errorOrTruncationOperation", "Insert"),
                                    new XAttribute("errorRowDisposition", "FailComponent"),
                                    new XAttribute("hasSideEffects", "true"),
                                    new XAttribute("name", "OLE DB Destination Input"));

            inputs.Add(data_Input);

            var data_InputColumns = new XElement("inputColumns");
            var data_ExternalMetadataColumns = new XElement("externalMetadataColumns",
                                                    new XAttribute("isUsed", "True"));
            var error_InputColumns = new XElement("inputColumns");



            var entityMetadata = Service.GetEntityMetadata(entityName);

            var EntAAIList_Src = CRMHelper.GetFullFieldList(Service, entityMetadata);

            foreach (var EntAAI in EntAAIList_Src)
            {
                GenerateXML_DataFlow_Component_InputColumn(data_InputColumns, EntAAI, data_path, lineage_path);
            }


            var EntAAIList_Stg = CRMHelper.GetFullFieldList(Service, entityMetadata, true);

            foreach (var EntAAI in EntAAIList_Stg)
            {
                GenerateXML_DataFlow_Component_externalMetadataColumn(data_ExternalMetadataColumns, EntAAI, data_path);
            }

            data_Input.Add(data_InputColumns);
            data_Input.Add(data_ExternalMetadataColumns);
        }
        private void GenerateXML_DataFlow_Component_Staging_Output(XElement component, string entityName, int dataFlowNumber)
        {
            var path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Destination.Outputs[OLE DB Destination Error Output]";
            var input_path = $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Destination.Inputs[OLE DB Destination Input]";

            XElement outputs = new XElement("outputs",
                new XElement("output",
                    new XAttribute("refId", path),
                    new XAttribute("exclusionGroup", "1"),
                    new XAttribute("isErrorOut", "true"),
                    new XAttribute("name", "OLE DB Destination Error Output"),
                    new XAttribute("synchronousInputId", input_path),
                    new XElement("outputColumns"),
                    new XElement("externalMetadataColumns")
                )
            );

            component.Add(outputs);

            var outputColumns = outputs.Element("output").Element("outputColumns");

            GenerateXML_DataFlow_Component_AddErrorOutputColumns(outputColumns, path);
        }

        private void GenerateXML_DataFlow_Component_AddErrorOutputColumns(XElement outputColumns, string path)
        {
            outputColumns.Add(
                        new XElement("outputColumn",
                            new XAttribute("refId", $"{path}.Columns[ErrorCode]"),
                            new XAttribute("dataType", "i4"),
                            new XAttribute("lineageId", $"{path}.Columns[ErrorCode]"),
                            new XAttribute("name", "ErrorCode"),
                            new XAttribute("specialFlags", "1")
                        ),
                        new XElement("outputColumn",
                            new XAttribute("refId", $"{path}.Columns[ErrorColumn]"),
                            new XAttribute("dataType", "i4"),
                            new XAttribute("lineageId", $"{path}.Columns[ErrorColumn]"),
                            new XAttribute("name", "ErrorColumn"),
                            new XAttribute("specialFlags", "2")
                        ));
        }

        private void GenerateXML_DataFlow_Component_Connections(XElement component, string entityName, int dataFlowNumber)
        {
            var connections = new XElement("connections");

            //@Chris - need to pass in database connection info
            if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBDestination")
            {
                connections.Add(new XElement("connection",
                                        new XAttribute("refId", $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Destination.Connections[OleDbConnection]"),
                                        new XAttribute("connectionManagerID", StagingDBId + ":external"),
                                        new XAttribute("connectionManagerRefId", "Project.ConnectionManagers[DESKTOP-C5HN73M_SQLEXPRESS.Staging_DMT]"),
                                        new XAttribute("name", "OleDbConnection")));
            }
            else if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBSource")
            {
                connections.Add(new XElement("connection",
                                        new XAttribute("refId", $"Package\\Data Flow Task_{dataFlowNumber}\\OLE DB Source.Connections[OleDbConnection]"),
                                        new XAttribute("connectionManagerID", SourceDBId + ":external"),
                                        new XAttribute("connectionManagerRefId", "Project.ConnectionManagers[DESKTOP-C5HN73M_SQLEXPRESS.AdventureWorks2022]"),
                                        new XAttribute("name", "OleDbConnection")));
            }
            else
            {
                var a = 1;
                //bad - add error handling?
            }

            component.Add(connections);
        }

        private void GenerateXML_DataFlow_Component_Properties(XElement component, string entityName)
        {
            var properties = new XElement("properties",
                                new XElement("property",
                                    new XAttribute("dataType", "System.Int32"),
                                    new XAttribute("description", "The number of seconds before a command times out.  A value of 0 indicates an infinite time-out."),
                                    new XAttribute("name", "CommandTimeout"),
                                    0
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "Specifies the variable that contains the name of the database object used to open a rowset."),
                                    new XAttribute("name", "OpenRowsetVariable")
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "The SQL command to be executed."),
                                    new XAttribute("name", "SqlCommand"),
                                    new XAttribute("UITypeEditor", "Microsoft.DataTransformationServices.Controls.ModalMultilineStringEditor")
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Int32"),
                                    new XAttribute("description", "Specifies the column code page to use when code page information is unavailable from the data source."),
                                    new XAttribute("name", "DefaultCodePage"),
                                    1252
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Boolean"),
                                    new XAttribute("description", "Forces the use of the DefaultCodePage property value when describing character data."),
                                    new XAttribute("name", "AlwaysUseDefaultCodePage"),
                                    false
                                )
            );

            if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBDestination")
            {
                properties.Add(new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "Specifies the name of the database object used to open a rowset."),
                                    new XAttribute("name", "OpenRowset"),
                                    $"[dbo].[{entityName}]"
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Int32"),
                                    new XAttribute("description", "Specifies the mode used to access the database."),
                                    new XAttribute("name", "AccessMode"),
                                    new XAttribute("typeConverter", "AccessMode"),
                                    3
                                ), 
                                new XElement("property",
                                    new XAttribute("dataType", "System.Boolean"),
                                    new XAttribute("description", "Indicates whether the values supplied for identity columns will be copied to the destination. If false, values for identity columns will be auto-generated at the destination. Applies only if fast load is turned on."),
                                    new XAttribute("name", "FastLoadKeepIdentity"),
                                    false
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Boolean"),
                                    new XAttribute("description", "Indicates whether the columns containing null will have null inserted in the destination. If false, columns containing null will have their default values inserted at the destination. Applies only if fast load is turned on."),
                                    new XAttribute("name", "FastLoadKeepNulls"),
                                    false
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "Specifies options to be used with fast load.  Applies only if fast load is turned on."),
                                    new XAttribute("name", "FastLoadOptions"),
                                    "TABLOCK,CHECK_CONSTRAINTS"
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Int32"),
                                    new XAttribute("description", "Specifies when commits are issued during data insertion.  A value of 0 specifies that one commit will be issued at the end of data insertion.  Applies only if fast load is turned on."),
                                    new XAttribute("name", "FastLoadMaxInsertCommitSize"),
                                    2147483647
                                )
                );
            }
            else if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBSource")
            {
                properties.Add(new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "Specifies the name of the database object used to open a rowset."),
                                    new XAttribute("name", "OpenRowset"),
                                    $"[dmt].[{entityName}]"
                                ),
                                new XElement("property",
                                    new XAttribute("dataType", "System.Int32"),
                                    new XAttribute("description", "Specifies the mode used to access the database."),
                                    new XAttribute("name", "AccessMode"),
                                    new XAttribute("typeConverter", "AccessMode"),
                                    0
                                ), 
                                new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "The mappings between the parameters in the SQL command and variables."),
                                    new XAttribute("name", "ParameterMapping")
                                ), 
                                new XElement("property",
                                    new XAttribute("dataType", "System.String"),
                                    new XAttribute("description", "The variable that contains the SQL command to be executed."),
                                    new XAttribute("name", "SqlCommandVariable")
                                )
                );
            }
            else
            {
                var a = 1;
                //bad - add error handling?
            }

            component.Add(properties);


        }

        /// <summary>
        /// This is used to link DTS:Executable>>DTS:Executables in the ControlFlow level of the SSIS package as opposed to the DataFlow level.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="fromExecutable_Name"></param>
        /// <param name="toExecutable_Name"></param>
        /// <param name="constraintNumber"></param>
        private void GenerateXML_PrecedenceConstraint(XDocument package ,string fromExecutable_Name, string toExecutable_Name, int constraintNumber) //@tim - how can I write back out to constraint number?
        {
            //@Testing - tested successfully by adding output to RunAll package

            var precedenceConstraints = package.Element(DTS + "Executable").Element(DTS + "PrecedenceConstraints");

            precedenceConstraints.Add(new XElement(DTS + "PrecedenceConstraint",
                                        new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                                        new XAttribute(DTS + "From", "Package\\" + fromExecutable_Name),
                                        new XAttribute(DTS + "To", "Package\\" + toExecutable_Name),
                                        new XAttribute(DTS + "LogicalAnd", "True"),
                                        new XAttribute(DTS + "ObjectName", "Constraint_" + constraintNumber)
                                        ));
        }

        private string DefineGuid(string guidBase, int guidEnd, bool addCurlyBrackets = false)
        {
            int guidEndLength = guidEnd.ToString().Length;

            var guidBaseTrimmed = guidBase.Substring(0, 36 - guidEndLength);

            var guidComplete = guidBaseTrimmed + guidEnd;

            if(addCurlyBrackets)
            {
                guidComplete = "{" + guidComplete + "}";
            }

            return guidComplete;
        }

        private void SavePackage(XDocument package, string entityName, string packageLocation)
        {
            package.Save(packageLocation + "\\" + entityName + ".dtsx");
        }


        public void addSSISPackageToProject(string entityName)
        {
            var project = GetS2SProject();
            var SSISPackageName = entityName.ToLower() + ".dtsx";

            XNamespace SSIS = "www.microsoft.com/SqlServer/SSIS";

            var SSISProject = project.Element("Project").Element("DeploymentModelSpecificContent").Element("Manifest").Element(SSIS + "Project");
            var SSISPackages = SSISProject.Element(SSIS + "Packages");
            var existingPackage = SSISPackages.Elements(SSIS + "Package").Where(x => x.Attribute(SSIS + "Name").Value.ToLower() == SSISPackageName).FirstOrDefault();

            if (existingPackage == null)
            {
                SSISPackages.Add(new XElement(SSIS + "Package",
                                    new XAttribute(SSIS + "Name", SSISPackageName),
                                    new XAttribute(SSIS + "EntryPoint", "1")));

                var SSISPackageInfo = SSISProject.Element(SSIS + "DeploymentInfo").Element(SSIS + "PackageInfo");


                SSISPackageInfo.Add(new XElement(SSIS + "PackageMetaData",
                    new XAttribute(SSIS + "Name", SSISPackageName),
                    new XElement(SSIS + "Properties",
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "ID"), GenerateNewXMLGuid()),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "Name"), entityName),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "VersionMajor"), "1"),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "VersionMinor"), "0"),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "VersionBuild"), "1"),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "VersionComments"), string.Empty),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "VersionGUID"), GenerateNewXMLGuid()),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "PackageFormatVersion"), "8"),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "Description"), string.Empty),
                        new XElement(SSIS + "Property",
                            new XAttribute(SSIS + "Name", "ProtectionLevel"), "1")
                    ),
                    new XElement(SSIS + "Parameters")
                ));

                try
                {
                    project.Save(mySettings.SourceToStagingLocationString + "\\SourceToStaging.dtproj");
                }
                catch (Exception e)
                {
                    MessageBox.Show("Unable to update \"SourceToStaging.dtproj\" to add the SSIS package to it.");
                }

            }
            
        }


        public XDocument GetS2SProject()
        {
            var path = mySettings.SourceToStagingLocationString + "\\SourceToStaging.dtproj";
            XDocument package = null;

            try
            {
                package = XDocument.Load(path);
            }
            catch (Exception e)
            {
                MessageBox.Show("Unable to find a file called \"SourceToStaging.dtproj\" at the Source To Staging SSIS Project Location. Please ensure you have the correct location set, e.g. C:\\Users\\Chris\\DMSolution\\SourceToStaging\n\n For more information please see the About button.");
            }

            return package;
        }

        public string GenerateNewXMLGuid()
        {
            return "{" + Guid.NewGuid().ToString().ToUpper() + "}";
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