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
            var ssisProjectLocation = mySettings.SourceToStagingLocationString;

            XMLGen = new XMLGeneration(Service, ssisProjectLocation);

            var package = new XDocument();

            var project = XMLGen.GetProjectFile();

            if (project != null)
            {
                SourceDBId = XMLGen.GetConmgr(ssisProjectLocation, "SourceDB").DTSID;
                StagingDBId = XMLGen.GetConmgr(ssisProjectLocation, "StagingDB").DTSID;

                if (SourceDBId != null && StagingDBId != null)
                {
                    XMLGen.GenerateXML_SSISPackageBase(package, entityName);
                    XMLGen.GenerateXML_Executable_SQLTask_BackupLoadInformation(package, entityName, mySettings.StagingDBSchema, StagingDBId);
                    XMLGen.GenerateXML_Executable_SQLTask_TruncateTable(package, entityName, mySettings.StagingDBSchema, StagingDBId);
                    GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(package, entityName);

                    var dataFlowTaskName = "DataFlowTask";

                    XMLGen.GenerateXML_Executable_DataFlow_Base(package, entityName, dataFlowTaskName);
                    GenerateXML_Executable_DataFlow_SimpleSourceToStaging_WithErrorHandling(package, entityName, dataFlowTaskName);
                    if (mySettings.UseDataTransforms)
                    {
                        GenerateXML_Executable_SQLTask_S2SDataTransforms(package, entityName);
                    }
                    GenerateXML_Executable_SQLTask_S2SReAddPK(package, entityName);
                    GenerateXML_Executable_SQLTask_S2SReAddIndexes(package, entityName);
                    XMLGen.GenerateXML_Executable_SQLTask_RestoreLoadInformation(package, entityName, mySettings.StagingDBSchema, StagingDBId);

                    //Constraints link Executables
                    int ConstraintNumber = 1;


                    XMLGen.GenerateXML_Executable_AddConstraint(package, $"Backup Load Information", $"Truncate {{{mySettings.StagingDBSchema}}}{{{entityName}}}", ConstraintNumber++);
                    XMLGen.GenerateXML_Executable_AddConstraint(package, $"Truncate {{{mySettings.StagingDBSchema}}}{{{entityName}}}", "Drop Indexes and PK", ConstraintNumber++);
                    XMLGen.GenerateXML_Executable_AddConstraint(package, "Drop Indexes and PK", dataFlowTaskName, ConstraintNumber++);
                    if (mySettings.UseDataTransforms)
                    {

                        XMLGen.GenerateXML_Executable_AddConstraint(package, dataFlowTaskName, "Data Transforms", ConstraintNumber++);
                        XMLGen.GenerateXML_Executable_AddConstraint(package, "Data Transforms", "ReAdd PK", ConstraintNumber++);
                    }
                    else
                    {
                        XMLGen.GenerateXML_Executable_AddConstraint(package, dataFlowTaskName, "ReAdd PK", ConstraintNumber++);
                    }
                    XMLGen.GenerateXML_Executable_AddConstraint(package, "ReAdd PK", "ReAdd Indexes", ConstraintNumber++);
                    XMLGen.GenerateXML_Executable_AddConstraint(package, "ReAdd Indexes", "Restore Load Information", ConstraintNumber++);

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


            var OLEDBSource_Name = $"DB - {{{mySettings.SourceDBSchema}}}{{{entityName}}}";
            var OLEDBDestination_Name = $"DB - {{{mySettings.StagingDBSchema}}}{{{entityName}}}";

            foreach (var entAAI in sourceFields)
            {
                entAAI.SSISLineage = OLEDBSource_Name;
            }

            XMLGen.GenerateXML_DataFlow_Component_OLEDBSource(components, entityName, mySettings.SourceDBSchema, dataFlowTaskName, sourceFields, "SourceDB", OLEDBSource_Name);
            XMLGen.GenerateXML_DataFlow_Component_OLEDBDestination(components, entityName, mySettings.StagingDBSchema, dataFlowTaskName, sourceFields, destinationFields, "StagingDB", OLEDBDestination_Name);


            var paths = executable.Element(DTS + "ObjectData").Element("pipeline").Element("paths");

            var pathStringBase = "Package\\" + dataFlowTaskName;


            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, OLEDBSource_Name, OLEDBDestination_Name);
        }



        private void GenerateXML_Executable_DataFlow_SimpleSourceToStaging_WithErrorHandling(XDocument package, string entityName, string dataFlowTaskName)
        {
            var sourceFields = CRMHelper.GetFullFieldList(Service, entityName);
            var destinationFields = CRMHelper.GetFullFieldList(Service, entityName, true);

            var executable = package.Element(DTS + "Executable").Element(DTS + "Executables").Elements(DTS + "Executable").Where(x => x.Attribute(DTS + "ObjectName").Value == dataFlowTaskName).FirstOrDefault();

            var components = executable.Element(DTS + "ObjectData").Element("pipeline").Element("components");

            var pathStringBase = "Package\\" + dataFlowTaskName;

            var OLEDBSource_Name = $"DB - {mySettings.SourceDBSchema} {entityName}";
            var OLEDBDestination_Name = $"DB - {mySettings.StagingDBSchema} {entityName}";

            foreach(var entAAI in sourceFields)
            {
                entAAI.SSISLineage = OLEDBSource_Name;
            }

            XMLGen.GenerateXML_DataFlow_Component_OLEDBSource(components, entityName, mySettings.SourceDBSchema, dataFlowTaskName, sourceFields, "SourceDB", OLEDBSource_Name);
            GenerateXML_DataFlow_Component_MultiCast_SourceToStaging(components, entityName, pathStringBase);

            //manually removing errordetails from sourcefield so ssis package doesn't try to add it to destination.
            sourceFields.RemoveAll(x => x.fieldName.ToLower() == "errordetails");
            XMLGen.GenerateXML_DataFlow_Component_OLEDBDestination(components, entityName, mySettings.StagingDBSchema, dataFlowTaskName, sourceFields, destinationFields, "StagingDB", OLEDBDestination_Name);

            GenerateXML_DataFlow_Component_ConditionalSplit_SourceToStaging(components, entityName, pathStringBase, OLEDBSource_Name);
            GenerateXML_DataFlow_Component_DerivedColumn_SourceToStaging(components, entityName, pathStringBase);

            //generate list of errorlogging columns
            var SourceErrorFields = new List<EntityAttribute_AdditionalInfo>();


            SourceErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, CRMHelper.GetEntityPK(Service, entityName) + "_Source", 255));
            SourceErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_IntegerField(entityName, "Source_System_Id"));
            SourceErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "ErrorDetails", 4000));
            SourceErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_UniqueIdentifier(entityName, "DynId"));

            foreach (var sf in SourceErrorFields)
            {
                sf.SSISLineage = OLEDBSource_Name;
            }

            var entAII = EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "Entity", 100);
            entAII.SSISLineage = "Derived Column";
            SourceErrorFields.Add(entAII);
            entAII = EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "LoadStep", 100);
            entAII.SSISLineage = "Derived Column";
            SourceErrorFields.Add(entAII);
            entAII = EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_DateTime(entityName, "ErrorDate");
            entAII.SSISLineage = "Derived Column";
            SourceErrorFields.Add(entAII);
            
            //generate list of errorlogging columns
            var TargetErrorFields = new List<EntityAttribute_AdditionalInfo>();

            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "Entity", 100));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "LoadStep", 100));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "id_Source", 255));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_IntegerField(entityName, "Source_System_Id"));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_DateTime(entityName, "ErrorDate"));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_String(entityName, "ErrorDetails", 4000));
            TargetErrorFields.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_Staging_UniqueIdentifier(entityName, "DynId"));

            //generate errorlogging oledbdestination
            var OLEDBDestination_ErrorLog_Name = $"DB - logging ErrorLog";
            XMLGen.GenerateXML_DataFlow_Component_OLEDBDestination(components, "ErrorLog", "logging", dataFlowTaskName, SourceErrorFields, TargetErrorFields, "StagingDB", OLEDBDestination_ErrorLog_Name);

            ///really bad code to find and replace one element becuase oledbdest doesn't work right///////////
            var entityPK = CRMHelper.GetEntityPK(Service, entityName);

            // Locate the specific Executable element by ObjectName attribute
            var dataflow_executable = package.Element(DTS + "Executable")
                                    ?.Elements(DTS + "Executables")
                                    ?.Elements(DTS + "Executable")
                                    .FirstOrDefault(x => x.Attribute(DTS + "ObjectName")?.Value == dataFlowTaskName);

            // Navigate through the structure to find the inputColumn element
            var OLEDBDestination_ErrorLog_Component = dataflow_executable.Element(DTS + "ObjectData")
                                                        ?.Element("pipeline")
                                                        ?.Element("components")
                                                        ?.Elements("component")
                                                        .FirstOrDefault(x => x.Attribute("name")?.Value == OLEDBDestination_ErrorLog_Name);


            var inputColumn = OLEDBDestination_ErrorLog_Component.Element("inputs")
                                        ?.Element("input")
                                        ?.Element("inputColumns")

                                        ?.Elements("inputColumn")
                                        .FirstOrDefault(x => x.Attribute("cachedName")?.Value == entityPK + "_Source");

            if (inputColumn != null)
            {
                var externalMetadataColumnIdAttr = inputColumn.Attribute("externalMetadataColumnId");

                if (externalMetadataColumnIdAttr != null && externalMetadataColumnIdAttr.Value.EndsWith($"[{entityPK}_Source]"))
                {
                    externalMetadataColumnIdAttr.Value = externalMetadataColumnIdAttr.Value.Replace($"[{entityPK}_Source]", "[id_Source]");
                }
            }
            ///bad code end////////////
            


            var paths = executable.Element(DTS + "ObjectData").Element("pipeline").Element("paths");

            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, OLEDBSource_Name, "Multicast", null, 1);
            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, "Multicast", OLEDBDestination_Name, 1);
            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, "Multicast", "Conditional Split", 2);
            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, "Conditional Split", "Derived Column");
            XMLGen.GenerateXML_DataFlow_path(paths, pathStringBase, "Derived Column", OLEDBDestination_ErrorLog_Name);
        }




        private void GenerateXML_DataFlow_Component_DerivedColumn_SourceToStaging(XElement components, string entityName, string path)
        {
            path = path + "\\Derived Column";

            // Define the XML structure
            XElement component = new XElement("component",
                new XAttribute("refId", path),
                new XAttribute("componentClassID", "Microsoft.DerivedColumn"),
                new XAttribute("contactInfo", "Derived Column;Microsoft Corporation; Microsoft SQL Server; (C) Microsoft Corporation; All Rights Reserved; http://www.microsoft.com/sql/support;0"),
                new XAttribute("description", "Creates new column values by applying expressions to transformation input columns. Create new columns or overwrite existing ones. For example, concatenate the values from the 'first name' and 'last name' column to make a 'full name' column."),
                new XAttribute("name", "Derived Column"),
                new XAttribute("usesDispositions", "true"),
                new XElement("inputs",
                    new XElement("input",
                        new XAttribute("refId", path + ".Inputs[Derived Column Input]"),
                        new XAttribute("description", "Input to the Derived Column Transformation"),
                        new XAttribute("name", "Derived Column Input"),
                        new XElement("externalMetadataColumns")
                    )
                ),
                new XElement("outputs",
                    new XElement("output",
                        new XAttribute("refId", path + ".Outputs[Derived Column Output]"),
                        new XAttribute("description", "Default Output of the Derived Column Transformation"),
                        new XAttribute("exclusionGroup", "1"),
                        new XAttribute("name", "Derived Column Output"),
                        new XAttribute("synchronousInputId", path + ".Inputs[Derived Column Input]"),
                        new XElement("outputColumns",
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Derived Column Output].Columns[Entity]"),
                                new XAttribute("dataType", "wstr"),
                                new XAttribute("errorOrTruncationOperation", "Computation"),
                                new XAttribute("errorRowDisposition", "FailComponent"),
                                new XAttribute("length", "100"),
                                new XAttribute("lineageId", path + ".Outputs[Derived Column Output].Columns[Entity]"),
                                new XAttribute("name", "Entity"),
                                new XAttribute("truncationRowDisposition", "FailComponent"),
                                new XElement("properties",
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Expression"),
                                        new XAttribute("name", "Expression"),
                                        $"\"{entityName}\""
                                    ),
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Friendly Expression"),
                                        new XAttribute("expressionType", "Notify"),
                                        new XAttribute("name", "FriendlyExpression"),
                                        $"\"{entityName}\""
                                    )
                                )
                            ),
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Derived Column Output].Columns[LoadStep]"),
                                new XAttribute("dataType", "wstr"),
                                new XAttribute("errorOrTruncationOperation", "Computation"),
                                new XAttribute("errorRowDisposition", "FailComponent"),
                                new XAttribute("length", "100"),
                                new XAttribute("lineageId", path + ".Outputs[Derived Column Output].Columns[LoadStep]"),
                                new XAttribute("name", "LoadStep"),
                                new XAttribute("truncationRowDisposition", "FailComponent"),
                                new XElement("properties",
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Expression"),
                                        new XAttribute("name", "Expression"),
                                        $"\"Source To Staging\""
                                    ),
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Friendly Expression"),
                                        new XAttribute("expressionType", "Notify"),
                                        new XAttribute("name", "FriendlyExpression"),
                                        $"\"Source To Staging\""
                                    )
                                )
                            ),
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Derived Column Output].Columns[ErrorDate]"),
                                new XAttribute("dataType", "dbTimeStamp"),
                                new XAttribute("errorOrTruncationOperation", "Computation"),
                                new XAttribute("errorRowDisposition", "FailComponent"),
                                new XAttribute("lineageId", path + ".Outputs[Derived Column Output].Columns[ErrorDate]"),
                                new XAttribute("name", "ErrorDate"),
                                new XAttribute("truncationRowDisposition", "FailComponent"),
                                new XElement("properties",
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Expression"),
                                        new XAttribute("name", "Expression"),
                                        "[GETDATE]()"
                                    ),
                                    new XElement("property",
                                        new XAttribute("containsID", "true"),
                                        new XAttribute("dataType", "System.String"),
                                        new XAttribute("description", "Derived Column Friendly Expression"),
                                        new XAttribute("expressionType", "Notify"),
                                        new XAttribute("name", "FriendlyExpression"),
                                        "GETDATE()"
                                    )
                                )
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    ),
                    new XElement("output",
                        new XAttribute("refId", path + ".Outputs[Derived Column Error Output]"),
                        new XAttribute("description", "Error Output of the Derived Column Transformation"),
                        new XAttribute("exclusionGroup", "1"),
                        new XAttribute("isErrorOut", "true"),
                        new XAttribute("name", "Derived Column Error Output"),
                        new XAttribute("synchronousInputId", path + ".Inputs[Derived Column Input]"),
                        new XElement("outputColumns",
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Derived Column Error Output].Columns[ErrorCode]"),
                                new XAttribute("dataType", "i4"),
                                new XAttribute("lineageId", path + ".Outputs[Derived Column Error Output].Columns[ErrorCode]"),
                                new XAttribute("name", "ErrorCode"),
                                new XAttribute("specialFlags", "1")
                            ),
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Derived Column Error Output].Columns[ErrorColumn]"),
                                new XAttribute("dataType", "i4"),
                                new XAttribute("lineageId", path + ".Outputs[Derived Column Error Output].Columns[ErrorColumn]"),
                                new XAttribute("name", "ErrorColumn"),
                                new XAttribute("specialFlags", "2")
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    )
                )
            );
            components.Add(component);

        }


        private void GenerateXML_DataFlow_Component_ConditionalSplit_SourceToStaging(XElement components, string entityName, string path, string OLEDBSource_Name)
        {
            path = path + "\\Conditional Split";

            XElement component = new XElement("component",
                new XAttribute("refId", path),
                new XAttribute("componentClassID", "Microsoft.ConditionalSplit"),
                new XAttribute("contactInfo", "Conditional Split;Microsoft Corporation; Microsoft SQL Server; (C) Microsoft Corporation; All Rights Reserved; http://www.microsoft.com/sql/support;0"),
                new XAttribute("description", "Routes data rows to different outputs depending on the content of the data. Use conditions (SSIS expressions) to specify which rows are routed. For example, separate records that need to be cleaned from those that are ready to be loaded or route only a subset of records."),
                new XAttribute("name", "Conditional Split"),
                new XAttribute("usesDispositions", "true"),
                new XElement("inputs",
                    new XElement("input",
                        new XAttribute("refId", path + ".Inputs[Conditional Split Input]"),
                        new XAttribute("description", "Input to the Conditional Split Transformation"),
                        new XAttribute("name", "Conditional Split Input"),
                        new XElement("inputColumns",
                            new XElement("inputColumn",
                                new XAttribute("refId", path + ".Inputs[Conditional Split Input].Columns[ErrorDetails]"),
                                new XAttribute("cachedDataType", "nText"),
                                new XAttribute("cachedName", "ErrorDetails"),
                                new XAttribute("lineageId", $"Package\\DataFlowTask\\{OLEDBSource_Name}.Outputs[{OLEDBSource_Name} Output].Columns[ErrorDetails]")
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    )
                ),
                new XElement("outputs",
                    new XElement("output",
                        new XAttribute("refId", path + ".Outputs[Conditional Split Output]"),
                        new XAttribute("description", "Output 1 of the Conditional Split Transformation"),
                        new XAttribute("errorOrTruncationOperation", "Computation"),
                        new XAttribute("errorRowDisposition", "FailComponent"),
                        new XAttribute("exclusionGroup", "1"),
                        new XAttribute("name", "Conditional Split Output"),
                        new XAttribute("synchronousInputId", path + ".Inputs[Conditional Split Input]"),
                        new XAttribute("truncationRowDisposition", "FailComponent"),
                        new XElement("properties",
                            new XElement("property",
                                new XAttribute("containsID", "true"),
                                new XAttribute("dataType", "System.String"),
                                new XAttribute("description", "Specifies the expression. This expression version uses lineage identifiers instead of column names."),
                                new XAttribute("name", "Expression"),
                                $"![ISNULL](#{{Package\\DataFlowTask\\{OLEDBSource_Name}.Outputs[{OLEDBSource_Name} Output].Columns[ErrorDetails]}})"
                            ),
                            new XElement("property",
                                new XAttribute("containsID", "true"),
                                new XAttribute("dataType", "System.String"),
                                new XAttribute("description", "Specifies the friendly version of the expression. This expression version uses column names."),
                                new XAttribute("expressionType", "Notify"),
                                new XAttribute("name", "FriendlyExpression"),
                                "!ISNULL(ErrorDetails)"
                            ),
                            new XElement("property",
                                new XAttribute("dataType", "System.Int32"),
                                new XAttribute("description", "Specifies the position of the condition in the list of conditions that the transformation evaluates. The evaluation order is from the lowest to the highest value."),
                                new XAttribute("name", "EvaluationOrder"),
                                "0"
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    ),
                    new XElement("output",
                        new XAttribute("refId", path + ".Outputs[Conditional Split Default Output]"),
                        new XAttribute("description", "Default Output of the Conditional Split Transformation"),
                        new XAttribute("exclusionGroup", "1"),
                        new XAttribute("name", "Conditional Split Default Output"),
                        new XAttribute("synchronousInputId", path + ".Inputs[Conditional Split Input]"),
                        new XElement("properties",
                            new XElement("property",
                                new XAttribute("dataType", "System.Boolean"),
                                new XAttribute("name", "IsDefaultOut"),
                                "true"
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    ),
                    new XElement("output",
                        new XAttribute("refId", path + ".Outputs[Conditional Split Error Output]"),
                        new XAttribute("description", "Error Output of the Conditional Split Transformation"),
                        new XAttribute("exclusionGroup", "1"),
                        new XAttribute("isErrorOut", "true"),
                        new XAttribute("name", "Conditional Split Error Output"),
                        new XAttribute("synchronousInputId", path + ".Inputs[Conditional Split Input]"),
                        new XElement("outputColumns",
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Conditional Split Error Output].Columns[ErrorCode]"),
                                new XAttribute("dataType", "i4"),
                                new XAttribute("lineageId", path + ".Outputs[Conditional Split Error Output].Columns[ErrorCode]"),
                                new XAttribute("name", "ErrorCode"),
                                new XAttribute("specialFlags", "1")
                            ),
                            new XElement("outputColumn",
                                new XAttribute("refId", path + ".Outputs[Conditional Split Error Output].Columns[ErrorColumn]"),
                                new XAttribute("dataType", "i4"),
                                new XAttribute("lineageId", path + ".Outputs[Conditional Split Error Output].Columns[ErrorColumn]"),
                                new XAttribute("name", "ErrorColumn"),
                                new XAttribute("specialFlags", "2")
                            )
                        ),
                        new XElement("externalMetadataColumns")
                    )
                )
            );

            components.Add(component);
        }

        private void GenerateXML_DataFlow_Component_MultiCast_SourceToStaging(XElement components, string entityName, string path)
        {
            path = path + "\\Multicast";

            XElement componentElement = new XElement("component",
                        new XAttribute("refId", path),
                        new XAttribute("componentClassID", "Microsoft.Multicast"),
                        new XAttribute("contactInfo", "Multicast;Microsoft Corporation; Microsoft SQL Server; (C) Microsoft Corporation; All Rights Reserved; http://www.microsoft.com/sql/support;0"),
                        new XAttribute("description", "Distributes every input row to every row in one or more outputs. For example, branch your data flow to make a copy of data so that some values can be masked before sharing with external partners."),
                        new XAttribute("name", "Multicast"),
                        new XElement("inputs",
                            new XElement("input",
                                new XAttribute("refId", $"{path}.Inputs[Multicast Input 1]"),
                                new XAttribute("name", "Multicast Input 1"),
                                new XElement("externalMetadataColumns")
                            )
                        ),
                        new XElement("outputs",
                            new XElement("output",
                                new XAttribute("refId", $"{path}.Outputs[Multicast Output 1]"),
                                new XAttribute("deleteOutputOnPathDetached", "true"),
                                new XAttribute("name", "Multicast Output 1"),
                                new XAttribute("synchronousInputId", $"{path}.Inputs[Multicast Input 1]"),
                                new XElement("externalMetadataColumns")
                            ),
                            new XElement("output",
                                new XAttribute("refId", $"{path}.Outputs[Multicast Output 2]"),
                                new XAttribute("deleteOutputOnPathDetached", "true"),
                                new XAttribute("name", "Multicast Output 2"),
                                new XAttribute("synchronousInputId", $"{path}.Inputs[Multicast Input 1]"),
                                new XElement("externalMetadataColumns")
                            ),
                            new XElement("output",
                                new XAttribute("refId", $"{path}.Outputs[Multicast Output 3]"),
                                new XAttribute("dangling", "true"),
                                new XAttribute("deleteOutputOnPathDetached", "true"),
                                new XAttribute("name", "Multicast Output 3"),
                                new XAttribute("synchronousInputId", $"{path}.Inputs[Multicast Input 1]"),
                                new XElement("externalMetadataColumns")
                            )
                        )
                    );
            components.Add(componentElement);
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