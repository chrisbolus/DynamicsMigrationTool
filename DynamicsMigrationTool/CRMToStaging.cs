using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DynamicsMigrationTool
{
    public class CRMToStagingGeneration
    {
        //See Diagram at the bottom for S2S package XML layout.

        private XMLGeneration XMLGen;
        public IOrganizationService Service { get; set; }
        public Settings mySettings { get; set; }
        XNamespace DTS = "www.microsoft.com/SqlServer/Dts";
        string CRMConnectionId = null;
        string StagingDBId = null;

        //this is the constructor
        public CRMToStagingGeneration(IOrganizationService service, Settings MySettings)
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
            var projectName = "CRMToStaging";
            var ssisProjectLocation = mySettings.CRMToStagingLocationString;

            XMLGen = new XMLGeneration(Service, ssisProjectLocation);

            var package = new XDocument();

            var project = XMLGen.GetProjectFile(projectName);

            if (project != null)
            {
                CRMConnectionId = XMLGen.GetConmgr(ssisProjectLocation, "Dynamics").DTSID;
                StagingDBId = XMLGen.GetConmgr(ssisProjectLocation, "StagingDB").DTSID;

                if (CRMConnectionId != null && StagingDBId != null)
                {
                    XMLGen.GenerateXML_SSISPackageBase(package, entityName);


                    var entAIIList = CRMHelper.GetFullFieldList(Service, entityName, false, true);

                    XMLGen.GenerateXML_Executable_SQLTask_DropAndCreateTable(package, entityName, mySettings.ImportSchema, StagingDBId, entAIIList);


                    var dataFlowTaskName = "Data Flow Task";

                    //gen dataflow
                    XMLGen.GenerateXML_Executable_DataFlow_Base(package, entityName, dataFlowTaskName);
                    GenerateXML_Executable_DataFlow_StagingToCRM(package, entityName, dataFlowTaskName);





                    //Constraints link Executables
                    int ConstraintNumber = 1;

                    XMLGen.GenerateXML_Executable_AddConstraint(package, $"Drop and Create {{{mySettings.ImportSchema}}}{{{entityName}}}", dataFlowTaskName, ConstraintNumber++);


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
                            MessageBox.Show("SSIS package failed to save. Check CRM To Staging SSIS Project Location is correct.");
                        }
                        XMLGen.addSSISPackageToProject(entityName, project);
                    }
                }
            }


        }



        private void GenerateXML_Executable_DataFlow_StagingToCRM(XDocument package, string entityName, string dataFlowTaskName)
        {
            var sourceFields = CRMHelper.GetFullFieldList(Service, entityName, false, true);

            var executable = package.Element(DTS + "Executable").Element(DTS + "Executables").Elements(DTS + "Executable").Where(x => x.Attribute(DTS + "ObjectName").Value == dataFlowTaskName).FirstOrDefault();

            var components = executable.Element(DTS + "ObjectData").Element("pipeline").Element("components");

            XMLGen.GenerateXML_DataFlow_Component_CRMSource(components, entityName, dataFlowTaskName, sourceFields, "Dynamics");
            XMLGen.GenerateXML_DataFlow_Component_OLEDBDestination(components, entityName, mySettings.ImportSchema, dataFlowTaskName, sourceFields, sourceFields, "StagingDB", "Dynamics CRM Source");

            var component_OLEDBDestination = components.Elements("component").Where(x => x.Attribute("componentClassID").Value == "Microsoft.OLEDBDestination").FirstOrDefault();

            component_OLEDBDestination.Add(new XAttribute("validateExternalMetadata", "false"));

            var paths = executable.Element(DTS + "ObjectData").Element("pipeline").Element("paths");

            var pathStringBase = "Package\\" + dataFlowTaskName;

            paths.Add(new XElement("path",
                            new XAttribute("refId", pathStringBase + ".Paths[Dynamics CRM Source Output]"),
                            new XAttribute("name", "Dynamics CRM Source Output"),
                            new XAttribute("startId", pathStringBase + "\\Dynamics CRM Source.Outputs[Dynamics CRM Source Output]"),
                            new XAttribute("endId", pathStringBase + "\\OLE DB Destination.Inputs[OLE DB Destination Input]")
                            ));


        }

    }
}
