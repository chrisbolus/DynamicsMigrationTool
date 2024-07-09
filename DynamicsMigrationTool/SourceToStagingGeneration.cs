using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using ScintillaNET;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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

        //this is the constructor
        public SourceToStagingGeneration(IOrganizationService service, Settings MySettings)
        {
            Service = service;
            mySettings = mySettings;
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

            package.Add(new XElement(DTS + "Executable",
                            new XAttribute(XNamespace.Xmlns + "DTS", DTS.ToString()),
                            new XAttribute(DTS + "ObjectName", entityName),
                            new XAttribute(DTS + "refId", "Package"),
                            new XAttribute(DTS + "DTSID", "{00000000-0000-0000-1000-000000000000}"),
                            new XAttribute(DTS + "ExecutableType", "Microsoft.Package"),
                            new XAttribute(DTS + "VersionBuild", "1"),
                            new XAttribute(DTS + "VersionGUID", "{00000000-0000-0000-1000-000000000000}"),
                            new XElement(DTS + "Property", "8",
                                new XAttribute (DTS + "Name", "PackageFormatVersion")
                            ),
                            new XElement(DTS + "Variables"),
                            new XElement(DTS + "Executables"),
                            new XElement(DTS + "PrecedenceConstraints")
                            ));


            GenerateXML_Executable_DataFlow_SimpleS2S(package, entityName, 1);

            try
            {
                SavePackage(package, entityName, packageLocation);
                MessageBox.Show("SSIS package created successfully.");
            }
            catch (Exception e)
            {
                MessageBox.Show("SSIS package failed to create.");
            }
        }

        private void GenerateXML_Executable_DataFlow_SimpleS2S(XDocument package, string entityName, int executableNumber)
        {
            var executables = package.Elements(DTS + "Executable").FirstOrDefault().Elements(DTS + "Executables").FirstOrDefault();

            var executable = new XElement(DTS + "Executable",
                                new XAttribute(DTS + "DTSID", DefineGuid("00000000-0000-0000-0000-010000000000", executableNumber, true)),
                                new XAttribute(DTS + "refId", "Package\\Data Flow Task_" + executableNumber),
                                new XAttribute(DTS + "CreationName", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "ExecutableType", "Microsoft.Pipeline"),
                                new XAttribute(DTS + "ObjectName", "Data Flow Task_" + executableNumber),
                                new XElement(DTS + "ObjectData",
                                    new XElement("pipeline",
                                        new XAttribute("version", "1"),
                                        new XElement("components"),
                                        new XElement("paths")
                                        )));

            executables.Add(executable);

            var components = executable.Elements(DTS + "ObjectData").FirstOrDefault().Elements("pipeline").FirstOrDefault().Elements("components").FirstOrDefault();

            GenerateXML_DataFlow_Component_Source(components, entityName, executableNumber);
            GenerateXML_DataFlow_Component_Staging(components, entityName, executableNumber);

            var paths = executable.Elements(DTS + "ObjectData").FirstOrDefault().Elements("pipeline").FirstOrDefault().Elements("paths").FirstOrDefault();

            var pathStringBase = "Package\\Data Flow Task_" + executableNumber;

            paths.Add(new XElement("path",
                            new XAttribute("refId", pathStringBase + ".Paths[OLE DB Source Output]"),
                            new XAttribute("name", "OLE DB Source Output"),
                            new XAttribute("startId", pathStringBase + "\\OLE DB Source.Outputs[OLE DB Source Output]"),
                            new XAttribute("endId", pathStringBase + "\\OLE DB Destination.Inputs[OLE DB Destination Input]")
                            ));
        }

        private void GenerateXML_DataFlow_Component_Source(XElement components, string entityName, int executableNumber)
        {
            var component = new XElement("component",
                                new XAttribute("refId", $"Package\\Data Flow Task_{executableNumber}\\OLE DB Source"),
                                new XAttribute("componentClassID", "Microsoft.OLEDBSource"),
                                new XAttribute("name", "OLE DB Source"),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "7")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_Properties(component, entityName);
            GenerateXML_DataFlow_Component_Connections(component, entityName, executableNumber);
            GenerateXML_DataFlow_Component_Source_Output(component, entityName, executableNumber);
        }


        private void GenerateXML_DataFlow_Component_Staging(XElement components, string entityName, int executableNumber)
        {
            var component = new XElement("component",
                                new XAttribute("refId", $"Package\\Data Flow Task_{executableNumber}\\OLE DB Destination"),
                                new XAttribute("componentClassID", "Microsoft.OLEDBDestination"),
                                new XAttribute("name", "OLE DB Destination"),
                                new XAttribute("usesDispositions", "true"),
                                new XAttribute("version", "4")
            );

            components.Add(component);

            GenerateXML_DataFlow_Component_Properties(component, entityName);
            GenerateXML_DataFlow_Component_Connections(component, entityName, executableNumber);
            GenerateXML_DataFlow_Component_Staging_Input(component, entityName, executableNumber);
            GenerateXML_DataFlow_Component_Staging_Output(component, entityName, executableNumber);
        }



        private void GenerateXML_DataFlow_Component_Source_Output(XElement component, string entityName, int executableNumber)
        {
            var outputs = new XElement("outputs");
            component.Add(outputs);

            var data_path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Source.Outputs[OLE DB Source Output]";
            var error_path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Source.Outputs[OLE DB Source Error Output]";


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

        private void GenerateXML_DataFlow_Component_Staging_Input(XElement component, string entityName, int executableNumber)
        {
            var inputs = new XElement("inputs");
            component.Add(inputs);

            var data_path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Destination.Inputs[OLE DB Destination Input]";
            var lineage_path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Source.Outputs[OLE DB Source Output]";




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
        private void GenerateXML_DataFlow_Component_Staging_Output(XElement component, string entityName, int executableNumber)
        {
            var path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Destination.Outputs[OLE DB Destination Error Output]";
            var input_path = $"Package\\Data Flow Task_{executableNumber}\\OLE DB Destination.Inputs[OLE DB Destination Input]";

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

            var outputColumns = outputs.Elements("output").FirstOrDefault().Elements("outputColumns").FirstOrDefault();

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

        private void GenerateXML_DataFlow_Component_Connections(XElement component, string entityName, int executableNumber)
        {
            var connections = new XElement("connections");

            //@Chris - need to pass in database connection info
            if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBDestination")
            {
                connections.Add(new XElement("connection",
                                        new XAttribute("refId", $"Package\\Data Flow Task_{executableNumber}\\OLE DB Destination.Connections[OleDbConnection]"),
                                        new XAttribute("connectionManagerID", "{00000000-0000-0000-0000-000000000000}:external"),
                                        new XAttribute("connectionManagerRefId", "Project.ConnectionManagers[DESKTOP-C5HN73M_SQLEXPRESS.Staging_DMT]"),
                                        new XAttribute("name", "OleDbConnection")));
            }
            else if (component.Attribute("componentClassID").Value == "Microsoft.OLEDBSource")
            {
                connections.Add(new XElement("connection",
                                        new XAttribute("refId", $"Package\\Data Flow Task_{executableNumber}\\OLE DB Source.Connections[OleDbConnection]"),
                                        new XAttribute("connectionManagerID", "{11111111-1111-1111-1111-111111111111}:external"),
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

            var precedenceConstraints = package.Elements(DTS + "Executable").FirstOrDefault().Elements(DTS + "PrecedenceConstraints").FirstOrDefault();

            precedenceConstraints.Add(new XElement(DTS + "PrecedenceConstraint",
                                        new XAttribute(DTS + "DTSID", DefineGuid("00000000-0000-0000-0000-200000000000", constraintNumber, true)),
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
//            //Source
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
//            //Staging
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