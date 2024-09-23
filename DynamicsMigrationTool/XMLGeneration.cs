using Microsoft.SqlServer.Management.Smo;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using ScintillaNET;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DynamicsMigrationTool
{
    internal class XMLGeneration
    {

        public IOrganizationService Service { get; set; }
        XNamespace DTS = "www.microsoft.com/SqlServer/Dts";
        string ssisProjectLocation;

        //this is the constructor
        public XMLGeneration(IOrganizationService service, string SSISProjectLocation)
        {
            Service = service;
            ssisProjectLocation = SSISProjectLocation;
        }

        //@tim do I need the update service and settings? In theory this will only be called by the tool during processes which will end before these are changed.
        internal void UpdateService(IOrganizationService newService)
        {
            Service = newService;
        }
        internal void UpdateSettings(string newSSISProjectLocation)
        {
            ssisProjectLocation = newSSISProjectLocation;
        }

        public class conMgr
        {
            public string DTSID;
            public string ObjectName;
        }


        public conMgr GetConmgr(string packageLocation, string ConnectionName)
        {
            var packagePath = packageLocation + $"\\{ConnectionName}.conmgr";
            var conMgrFile = new XDocument();
            var conmgr = new conMgr();

            Boolean goodPath = false;

            try
            {
                conMgrFile = XDocument.Load(packagePath);
                goodPath = true;
            }
            catch (Exception e)
            {
                MessageBox.Show($"Unable to find a file called {ConnectionName}.conmgr at {packageLocation}");
            }

            if (goodPath)
            {
                try
                {
                    conmgr.DTSID = conMgrFile.Element(DTS + "ConnectionManager").Attribute(DTS + "DTSID").Value;
                    conmgr.ObjectName = conMgrFile.Element(DTS + "ConnectionManager").Attribute(DTS + "ObjectName").Value;
                }
                catch (Exception e)
                {
                    MessageBox.Show($"{ConnectionName}.conmgr file layout is incorrect.");
                }
            }

            return conmgr;
        }

        public void GenerateXML_SSISPackageBase(XDocument package, string packageName)
        {
            var loggingGuid = GenerateNewXMLGuid();
            var loggingDB = "StagingDB";

            package.Add(new XElement(DTS + "Executable",
                            new XAttribute(XNamespace.Xmlns + "DTS", DTS.ToString()),
                            new XAttribute(DTS + "ObjectName", packageName),
                            new XAttribute(DTS + "refId", "Package"),
                            new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                            new XAttribute(DTS + "ExecutableType", "Microsoft.Package"),
                            new XAttribute(DTS + "CreationName", "Microsoft.Package"),
                            new XAttribute(DTS + "VersionBuild", "1"),
                            new XAttribute(DTS + "VersionGUID", GenerateNewXMLGuid()),
                            new XElement(DTS + "Property", "8",
                                new XAttribute(DTS + "Name", "PackageFormatVersion")
                            ),
                            new XElement(DTS + "LogProviders",
                                new XElement(DTS + "LogProvider",
                                new XAttribute(DTS + "ConfigString", GetConmgr(ssisProjectLocation, loggingDB).ObjectName),
                                new XAttribute(DTS + "CreationName", "Microsoft.LogProviderSQLServer"),
                                new XAttribute(DTS + "Description", "Writes log entries for events to a SQL Server database"),
                                new XAttribute(DTS + "DTSID", loggingGuid),
                                new XAttribute(DTS + "ObjectName", "SSIS log provider for SQL Server"),
                                new XElement(DTS + "ObjectData",
                                    new XElement("InnerObject") // InnerObject element
                                ))
                            ),
                            new XElement(DTS + "LoggingOptions",
                                new XAttribute(DTS + "FilterKind", "0"),
                                new XAttribute(DTS + "LoggingMode", "1"),
                                new XElement(DTS + "Property",
                                    new XAttribute(DTS + "DataType", "8"),
                                    new XAttribute(DTS + "Name", "EventFilter"),
                                    "5,7,OnError,13,OnPostExecute,12,OnPreExecute,13,OnQueryCancel,12,OnTaskFailed"
                                ),
                                CreateColumnFilter(DTS, "OnError"),
                                CreateColumnFilter(DTS, "OnPostExecute"),
                                CreateColumnFilter(DTS, "OnPreExecute"),
                                CreateColumnFilter(DTS, "OnQueryCancel"),
                                CreateColumnFilter(DTS, "OnTaskFailed"),
                                new XElement(DTS + "SelectedLogProviders",
                                    new XElement(DTS + "SelectedLogProvider",
                                        new XAttribute(DTS + "InstanceID", loggingGuid)
                                ))
                            ),
                            new XElement(DTS + "Variables"),
                            new XElement(DTS + "Executables"),
                            new XElement(DTS + "PrecedenceConstraints")
                            ));



        }


        static XElement CreateColumnFilter(XNamespace DTS, string eventName)
        {
            return new XElement(DTS + "Property",
                new XAttribute(DTS + "EventName", eventName),
                new XAttribute(DTS + "Name", "ColumnFilter"),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "Computer"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "Operator"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "SourceName"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "SourceID"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "ExecutionID"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "MessageText"), "-1"
                ),
                new XElement(DTS + "Property",
                    new XAttribute(DTS + "Name", "DataBytes"), "-1"
                )
            );
        }

        public void GenerateXML_Executable_SQLTask_DropAndCreateTable(XDocument package, string entityName, string schemaName, string SQLConnection, List<EntityAttribute_AdditionalInfo> EntAAIList)
        {
            var SQLString = $"IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[{schemaName}].{entityName}') AND type in (N'U')) \n" +
                $"DROP TABLE [{schemaName}].{entityName}\n\n" +
                $"CREATE TABLE [{schemaName}].{entityName} \n(";

            foreach (var EntAAI in EntAAIList)
            {
                SQLString = SQLString + $"{EntAAI.fieldName} {EntAAI.DBDataType} {(EntAAI.stgNotNull ? " NOT NULL" : "")},\n";
            }
            SQLString = SQLString + "\n)";

            GenerateXML_Executable_SQLTask_AddTask(package, $"Drop and Create {{{schemaName}}}{{{entityName}}}", SQLConnection, SQLString);
        }

        public void GenerateXML_Executable_SQLTask_TruncateTable(XDocument package, string entityName, string schemaName, string SQLConnection)
        {
            var SQLQuery = $"truncate table {schemaName}.{entityName}";

            GenerateXML_Executable_SQLTask_AddTask(package, $"Truncate {{{schemaName}}}{{{entityName}}}", SQLConnection, SQLQuery);
        }

        public void GenerateXML_Executable_SQLTask_DeleteErrorLogs(XDocument package, string entityName, string schemaName, string SQLConnection)
        {
            var SQLQuery = $"DELETE \r\n" +
                            $"FROM Logging.ErrorLog \r\n" +
                            $"WHERE LoadStep = 'Source To Staging' \r\n" +
                            $"AND Entity = '{entityName}'";

            GenerateXML_Executable_SQLTask_AddTask(package, $"Delete Source To Staging Error Logs for {entityName}", SQLConnection, SQLQuery);
        }

        public void GenerateXML_Executable_SQLTask_BackupLoadInformation(XDocument package, string entityName, string schemaName, string SQLConnection)
        {
            var entityPK = CRMHelper.GetEntityPK(Service, entityName);

            var SQLQuery = $"IF NOT EXISTS(SELECT * FROM SYS.TABLES AS T INNER JOIN SYS.schemas AS S ON T.schema_id = S.schema_id WHERE S.NAME = '{schemaName}' AND T.NAME = 'TMP_{entityName}')\r\n" +
                            $"SELECT {entityPK}_Source, Source_System_Id, DynId, Processing_Status, DateCreate, DateUpdate, DateDelete\r\n" +
                            $"INTO {schemaName}.TMP_{entityName}\r\n" +
                            $"FROM {schemaName}.{entityName}";

            GenerateXML_Executable_SQLTask_AddTask(package, $"Backup Load Information", SQLConnection, SQLQuery);
        }

        public void GenerateXML_Executable_SQLTask_RestoreLoadInformation(XDocument package, string entityName, string schemaName, string SQLConnection)
        {
            var entityPK = CRMHelper.GetEntityPK(Service, entityName);


            //have had to remove processing status as without delta flagging, records that have changed will still be flagged loaded
            var SQLQuery = $"UPDATE e\r\n" +
                            $"SET\te.DynId = t.DynId, " +
                            //$"e.Processing_Status = t.Processing_Status, " +
                            $"e.DateCreate = t.DateCreate, " +
                            $"e.DateUpdate = t.DateUpdate, " +
                            $"e.DateDelete = t.DateDelete\r\n" +
                            $"FROM {schemaName}.{entityName} AS e\r\n" +
                            $"INNER JOIN {schemaName}.TMP_{entityName} AS t ON e.{entityPK}_Source = t.{entityPK}_Source AND e.Source_System_Id = t.Source_System_Id\r\n" +
                            $"GO\r\n" +
                            $"DROP TABLE {schemaName}.TMP_{entityName}\r\n" +
                            $"GO";

            GenerateXML_Executable_SQLTask_AddTask(package, $"Restore Load Information", SQLConnection, SQLQuery);
        }

        public void GenerateXML_Executable_SQLTask_AddTask(XDocument package, string executableName, string SQLConnection, string SQLQuery)
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
                new XElement(DTS + "LoggingOptions",
                    new XAttribute("FilterKind", "0")),
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

        public void GenerateXML_Executable_AddConstraint(XDocument package, string executableName_from, string executableName_to, int ConstraintNumber)
        {
            var precedenceConstraints = package.Element(DTS + "Executable").Element(DTS + "PrecedenceConstraints");

            XElement precedenceConstraint = new XElement(DTS + "PrecedenceConstraint",
                new XAttribute(DTS + "From", $"Package\\{executableName_from}"),
                new XAttribute(DTS + "To", $"Package\\{executableName_to}"),
                new XAttribute(DTS + "refId", $"Package.PrecedenceConstraints[Constraint_{ConstraintNumber}]"),
                new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                new XAttribute(DTS + "LogicalAnd", "True"),
                new XAttribute(DTS + "ObjectName", "Constraint"),
                new XAttribute(DTS + "CreationName", string.Empty)
            );

            precedenceConstraints.Add(precedenceConstraint);
        }


        public void GenerateXML_DataFlow_Component_OLEDBSource(XElement components, string entityName, string schemaName, string dataFlowTaskName, List<EntityAttribute_AdditionalInfo> SourceFields, string conmgrName, string OLEDBSource_Name)
        {
            var path = $"Package\\{dataFlowTaskName}\\{OLEDBSource_Name}";

            var component = new XElement("component",
                                new XAttribute("refId", path),
                                new XAttribute("componentClassID", "Microsoft.OLEDBSource"),
                                new XAttribute("name", OLEDBSource_Name),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "7")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_OLEDB_Properties(component, entityName, schemaName);
            GenerateXML_DataFlow_Component_Connections(component, "OleDbConnection", path, conmgrName);
            GenerateXML_DataFlow_Component_OLEDBSource_Output(component, entityName, dataFlowTaskName, SourceFields, OLEDBSource_Name);
        }


        public void GenerateXML_DataFlow_Component_OLEDBDestination(XElement components, string entityName, string schemaName, string dataFlowTaskName, List<EntityAttribute_AdditionalInfo> SourceFields, List<EntityAttribute_AdditionalInfo> DestinationFields, string conmgrName, string OLEDBDestination_Name)
        {
            var path = $"Package\\{dataFlowTaskName}\\{OLEDBDestination_Name}";

            var component = new XElement("component",
                                new XAttribute("refId", path),
                                new XAttribute("componentClassID", "Microsoft.OLEDBDestination"),
                                new XAttribute("name", OLEDBDestination_Name),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "4")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_OLEDB_Properties(component, entityName, schemaName);
            GenerateXML_DataFlow_Component_Connections(component, "OleDbConnection", path, conmgrName);
            GenerateXML_DataFlow_Component_OLEDBDestination_Input(component, dataFlowTaskName, SourceFields, DestinationFields, OLEDBDestination_Name);
            GenerateXML_DataFlow_Component_OLEDBDestination_Output(component, dataFlowTaskName, OLEDBDestination_Name);
        }



        private void GenerateXML_DataFlow_Component_OLEDBSource_Output(XElement component, string entityName, string dataFlowTaskName, List<EntityAttribute_AdditionalInfo> SourceFields, string OLEDBSource_Name)
        {
            var outputs = new XElement("outputs");
            component.Add(outputs);

            var data_path = $"Package\\{dataFlowTaskName}\\{OLEDBSource_Name}.Outputs[{OLEDBSource_Name} Output]";
            var error_path = $"Package\\{dataFlowTaskName}\\{OLEDBSource_Name}.Outputs[{OLEDBSource_Name} Error Output]";


            var data_Output = new XElement("output",
                                    new XAttribute("refId", data_path),
                                    new XAttribute("name", $"{OLEDBSource_Name} Output"));

            var error_Output = new XElement("output",
                                    new XAttribute("refId", error_path),
                                    new XAttribute("isErrorOut", "true"),
                                    new XAttribute("name", $"{OLEDBSource_Name} Error Output"));

            outputs.Add(data_Output);
            outputs.Add(error_Output);

            var data_OutputColumns = new XElement("outputColumns");
            var data_ExternalMetadataColumns = new XElement("externalMetadataColumns",
                                                    new XAttribute("isUsed", "True"));
            var error_OutputColumns = new XElement("outputColumns");

            GenerateXML_DataFlow_Component_AddErrorOutputColumns(error_OutputColumns, error_path);

            foreach (var EntAAI in SourceFields)
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
                externalMetadataColumn.Add(new XAttribute("precision", "23"));
                externalMetadataColumn.Add(new XAttribute("scale", "10"));
            }
            data_ExternalMetadataColumns.Add(externalMetadataColumn);
        }

        private void GenerateXML_DataFlow_Component_OLEDBDestination_Input(XElement component, string dataFlowTaskName, List<EntityAttribute_AdditionalInfo> SourceFields, List<EntityAttribute_AdditionalInfo> DestinationFields, string OLEDBDestination_Name)
        {
            var inputs = new XElement("inputs");
            component.Add(inputs);

            var data_path = $"Package\\{dataFlowTaskName}\\{OLEDBDestination_Name}.Inputs[{OLEDBDestination_Name} Input]";



            var data_Input = new XElement("input",
                                    new XAttribute("refId", data_path),
                                    new XAttribute("errorOrTruncationOperation", "Insert"),
                                    new XAttribute("errorRowDisposition", "FailComponent"),
                                    new XAttribute("hasSideEffects", "true"),
                                    new XAttribute("name", $"{OLEDBDestination_Name} Input"));

            inputs.Add(data_Input);

            var data_InputColumns = new XElement("inputColumns");
            var data_ExternalMetadataColumns = new XElement("externalMetadataColumns",
                                                    new XAttribute("isUsed", "True"));
            var error_InputColumns = new XElement("inputColumns");



            foreach (var EntAAI in SourceFields)
            {

                var lineage_path = $"Package\\{dataFlowTaskName}\\{EntAAI.SSISLineage}.Outputs[{EntAAI.SSISLineage} Output]";

                GenerateXML_DataFlow_Component_InputColumn(data_InputColumns, EntAAI, data_path, lineage_path);
            }


            foreach (var EntAAI in DestinationFields)
            {
                //handling situation where we have a decimal from CRM source, but the DB needs it to be a numeric for this step.
                var EntAAI_Modified = EntAAI;

                if(EntAAI_Modified.SSISDataType == "decimal")
                {
                    EntAAI_Modified.SSISDataType = "numeric";
                }
                GenerateXML_DataFlow_Component_externalMetadataColumn(data_ExternalMetadataColumns, EntAAI_Modified, data_path);
            }

            data_Input.Add(data_InputColumns);
            data_Input.Add(data_ExternalMetadataColumns);
        }
        private void GenerateXML_DataFlow_Component_OLEDBDestination_Output(XElement component, string dataFlowTaskName, string OLEDBDestination_Name)
        {
            var path = $"Package\\{dataFlowTaskName}\\{OLEDBDestination_Name}.Outputs[{OLEDBDestination_Name} Error Output]";
            var input_path = $"Package\\{dataFlowTaskName}\\{OLEDBDestination_Name}.Inputs[{OLEDBDestination_Name} Input]";

            XElement outputs = new XElement("outputs",
                new XElement("output",
                    new XAttribute("refId", path),
                    new XAttribute("exclusionGroup", "1"),
                    new XAttribute("isErrorOut", "true"),
                    new XAttribute("name", $"{OLEDBDestination_Name} Error Output"),
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


        private void GenerateXML_DataFlow_Component_Connections(XElement component, string connectionType, string path, string conmgrName)
        {
            var connections = new XElement("connections");

            connections.Add(new XElement("connection",
                                        new XAttribute("refId", $"{path}.Connections[{connectionType}]"),
                                        new XAttribute("connectionManagerID", GetConmgr(ssisProjectLocation, conmgrName).DTSID + ":external"),
                                        new XAttribute("connectionManagerRefId", $"Project.ConnectionManagers[{GetConmgr(ssisProjectLocation, conmgrName).ObjectName}]"),
                                        new XAttribute("name", connectionType)));

            component.Add(connections);
        }

        private void GenerateXML_DataFlow_Component_OLEDB_Properties(XElement component, string entityName, string schemaName)
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
                                    $"[{schemaName}].[{entityName}]"
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
                                    $"[{schemaName}].[{entityName}]"
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


        private void GenerateXML_DataFlow_Component_CRMSource_Properties(XElement component, string entityName)
        {
            var properties = new XElement("properties",
                new XElement("property",
                    new XAttribute("dataType", "System.Int32"),
                    new XAttribute("description", "Specifies the type of source data from Microsoft Dynamics CRM."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "SourceType"),
                    new XAttribute("typeConverter", "KingswaySoft.IntegrationToolkit.DynamicsCrm.SourceType"),
                    0),
                new XElement("property",
                    new XAttribute("dataType", "System.String"),
                    new XAttribute("description", "Dynamics CRM entity to retrieve data from."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "SourceEntity"),
                    entityName),
                new XElement("property",
                    new XAttribute("dataType", "System.String"),
                    new XAttribute("description", "FetchXML statement."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "FetchXML")),
                new XElement("property",
                    new XAttribute("dataType", "System.Int32"),
                    new XAttribute("description", "Specifies the batch size of the query."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "BatchSize"),
                    2000),
                new XElement("property",
                    new XAttribute("dataType", "System.Int32"),
                    new XAttribute("description", "Specify the maximum number of rows to be returned from the Dynamics 365/CRM server (Specify 0 to read all satisfying records)."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "MaxRowsReturned"),
                    0),
                new XElement("property",
                    new XAttribute("dataType", "System.Null"),
                    new XAttribute("description", "The caller to be used for CRM impersonation when reading from CRM server."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "ImpersonateAs")),
                new XElement("property",
                    new XAttribute("dataType", "System.Null"),
                    new XAttribute("description", "Specifies the output timezone for CRM datetime fields."),
                    new XAttribute("expressionType", "Notify"),
                    new XAttribute("name", "OutputTimezone"),
                    new XAttribute("typeConverter", "KingswaySoft.IntegrationToolkit.Common.DescriptiveEnumTypeConverter`1[[KingswaySoft.IntegrationToolkit.DynamicsCrm.SourceOutputTimezone")),
                new XElement("property",
                    new XAttribute("dataType", "System.String"),
                    new XAttribute("name", "UserComponentTypeName"),
                    "KingswaySoft.IntegrationToolkit.DynamicsCrm.CrmSourceComponent")
                );

            component.Add(properties);
        }


        public void GenerateXML_DataFlow_path(XElement paths, string path, string sourceName, string targetName, int? sourceOutputNumber = null, int? targetInputNumber = null)
        {
            var targetInputNumberString = "";
            var outputNumberString = "";

            if (targetInputNumber != null)
            {
                targetInputNumberString = " " + targetInputNumber.ToString();
            }

            if (sourceOutputNumber != null)
            {
                outputNumberString = " " + sourceOutputNumber.ToString();
            }

            paths.Add(new XElement("path",
                            new XAttribute("refId", $"{path}.Paths[{sourceName} Output{outputNumberString}]"),
                            new XAttribute("name", $"{sourceName} Output{outputNumberString}"),
                            new XAttribute("startId", $"{path}\\{sourceName}.Outputs[{sourceName} Output{outputNumberString}]"),
                            new XAttribute("endId", $"{path}\\{targetName}.Inputs[{targetName} Input{targetInputNumberString}]")
                            ));
        }




        public void SavePackage(XDocument package, string packageName, string packageLocation)
        {
            package.Save(packageLocation + "\\" + packageName + ".dtsx");
        }

        public XDocument createRunAll(XDocument project)
        {
            var packageName = "_RunAll";
            var package = new XDocument();
            GenerateXML_SSISPackageBase(package, packageName);

            var executables = package.Element(DTS + "Executable").Element(DTS + "Executables");


            XNamespace SSIS = "www.microsoft.com/SqlServer/SSIS";

            var SSISProject = project.Element("Project").Element("DeploymentModelSpecificContent").Element("Manifest").Element(SSIS + "Project");
            var SSISPackages = SSISProject.Element(SSIS + "Packages");
            var existingPackages = SSISPackages.Elements(SSIS + "Package");
            int contraintNumber = 1;
            var previousPackageName = "";

            foreach (var existingPackage in existingPackages)
            {
                var existingPackageName = existingPackage.Attribute(SSIS + "Name").Value.Replace(".dtsx", "");

                if (!existingPackageName.ToLower().Contains("runall"))
                {
                    XElement executable = new XElement(DTS + "Executable",
                        new XAttribute(DTS + "refId", $"Package\\Execute Package - {existingPackageName}"),
                        new XAttribute(DTS + "CreationName", "Microsoft.ExecutePackageTask"),
                        new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                        new XAttribute(DTS + "ExecutableType", "Microsoft.ExecutePackageTask"),
                        new XAttribute(DTS + "LocaleID", "-1"),
                        new XAttribute(DTS + "ObjectName", $"Execute Package - {existingPackageName}"),
                        new XElement(DTS + "Variables"),
                        new XElement(DTS + "ObjectData",
                            new XElement("ExecutePackageTask",
                                new XElement("UseProjectReference", "True"),
                                new XElement("PackageName", $"{existingPackageName}.dtsx")
                            )
                        )
                    );

                    executables.Add(executable);

                    if (previousPackageName != "")
                    {
                        GenerateXML_Executable_AddConstraint(package, $"Execute Package - {previousPackageName}", $"Execute Package - {existingPackageName}", contraintNumber);
                        contraintNumber++;
                    }

                    previousPackageName = existingPackageName;
                }
                
            }


            return package;

        }

        


        public void addSSISPackageToProject(string packageName, XDocument project)
        {
            var SSISPackageName = packageName.ToLower() + ".dtsx";

            XNamespace SSIS = "www.microsoft.com/SqlServer/SSIS";

            var projectName = project.Element("Project").Element("DeploymentModelSpecificContent").Element("Manifest").Element(SSIS + "Project").Element(SSIS + "Properties").Elements(SSIS + "Property").Where(x => x.Attribute(SSIS + "Name").Value == "Name").FirstOrDefault().Value;
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
                            new XAttribute(SSIS + "Name", "Name"), packageName),
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
                    project.Save(ssisProjectLocation + $"\\{projectName}.dtproj");
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Unable to update \"{projectName}.dtproj\" to add the SSIS package to it.");
                }

            }

        }


        public string GenerateNewXMLGuid()
        {
            return "{" + Guid.NewGuid().ToString().ToUpper() + "}";
        }


        public XDocument GetProjectFile()
        {
            var path = "";

            try
            {
                var files = Directory.GetFiles(ssisProjectLocation, "*.dtproj");

                if (files == null)
                {
                    MessageBox.Show($"Unable to find a file called \"*.dtproj\" at the {ssisProjectLocation}. Please ensure you have the correct location set, e.g. C:\\Users\\Chris\\DMSolution\\[projectName]\n\n For more information please see the About button.");
                }
                else if (files.Count() > 1)
                {
                    MessageBox.Show($"Found multiple files called \"*.dtproj\" at the {ssisProjectLocation}. Please move/remove additional .dtproj files.");
                }
                else
                {
                    path = files.FirstOrDefault();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"Unable to find a file called \"*.dtproj\" at the {ssisProjectLocation}. Please ensure you have the correct location set, e.g. C:\\Users\\Chris\\DMSolution\\[projectName]\n\n For more information please see the About button.");
            }



            XDocument package = null;

            if (path != "")
            { 
                try
                {
                    package = XDocument.Load(path);
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Unable to find {path}. Please ensure you have the correct location set.");
                }
            }

            return package;
        }

        internal void GenerateXML_Executable_DataFlow_Base(XDocument package, string entityName, string dataFlowTaskName)
        {
            var executables = package.Element(DTS + "Executable").Element(DTS + "Executables");

            var executable = new XElement(DTS + "Executable",
                                new XAttribute(DTS + "DTSID", GenerateNewXMLGuid()),
                                new XAttribute(DTS + "refId", "Package\\" + dataFlowTaskName),
                                new XAttribute(DTS + "ObjectName", dataFlowTaskName),
                                new XAttribute(DTS + "CreationName", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "ExecutableType", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "LocaleID", "-1"),
                                new XElement(DTS + "LoggingOptions",
                                    new XAttribute("FilterKind", "0")),
                                new XElement(DTS + "ObjectData",
                                    new XElement("pipeline",
                                        new XAttribute("version", "1"),
                                        new XElement("components"),
                                        new XElement("paths")
                                        )));

            executables.Add(executable);

        }



        internal void GenerateXML_DataFlow_Component_CRMSource(XElement components, string entityName, string dataFlowTaskName, List<EntityAttribute_AdditionalInfo> crmFields, string conmgrName, string CRMSource_Name)
        {

            var path = $"Package\\{dataFlowTaskName}\\{CRMSource_Name}";


            XElement component = new XElement("component",
                new XAttribute("refId", path),
                new XAttribute("componentClassID", "Microsoft.ManagedComponentHost"),
                new XAttribute("description", "Extract data from Microsoft Dynamics 365 CE/CRM."),
                new XAttribute("name", CRMSource_Name),
                new XAttribute("usesDispositions", "true"),
                new XAttribute("version", "3")                
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_CRMSource_Properties(component, entityName);
            GenerateXML_DataFlow_Component_Connections(component, "DynamicsCRM", path, conmgrName);
            GenerateXML_DataFlow_Component_CRMSource_Output(component, entityName, path, crmFields, CRMSource_Name);
        }

        private void GenerateXML_DataFlow_Component_CRMSource_Output(XElement component, string entityName, string path, List<EntityAttribute_AdditionalInfo> crmFields, string CRMSource_Name)
        {
            var outputs = new XElement("outputs"
                            );

            component.Add(outputs);


            var output = new XElement("output",
                                new XAttribute("refId", $"{path}.Outputs[{CRMSource_Name} Output]"),
                                new XAttribute("name", $"{CRMSource_Name} Output")
                            );

            outputs.Add(output);

            path = $"{path}.Outputs[{CRMSource_Name} Output]";

            var outputColumns = new XElement("outputColumns");
            var externalMetadataColumns = new XElement("externalMetadataColumns", new XAttribute("isUsed", "True"));
            output.Add(outputColumns);
            output.Add(externalMetadataColumns);

            foreach (var field in crmFields)
            {

                GenerateXML_DataFlow_Component_OutputColumn(outputColumns, field, path);
                GenerateXML_DataFlow_Component_externalMetadataColumn(externalMetadataColumns, field, path);


            }
        }
    }
}
