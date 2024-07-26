using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using System;
using System.Linq;
using System.Web.Services.Description;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DynamicsMigrationTool
{
    public class StagingToCRMGeneration
    {
        //See Diagram at the bottom for S2C package XML layout. --make this

        private XMLGeneration XMLGen;

        public IOrganizationService Service { get; set; }
        public Settings mySettings { get; set; }
        string SourceDBId = null;
        string StagingDBId = null;

        //this is the constructor
        public StagingToCRMGeneration(IOrganizationService service, Settings MySettings)
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
            var packageLocation = mySettings.StagingToCRMLocationString;

            XMLGen = new XMLGeneration(Service, packageLocation);

            var package = new XDocument();

            var project = GetStagingToCRMProject();

            //if (project != null)
            //{
            //    SourceDBId = XMLGen.GetConmgrId(packageLocation, "SourceDB");
            //    StagingDBId = GetConmgrId(packageLocation, "StagingDB");

            //    if (SourceDBId != null && StagingDBId != null)
            //    {
            //        GenerateXML_SSISPackageBase(package, entityName);
            //        GenerateXML_Executable_SQLTask_S2STruncateStagingTable(package, entityName);
            //        GenerateXML_Executable_SQLTask_S2SDropIndexesAndPK(package, entityName);
            //        GenerateXML_Executable_DataFlow_SimpleS2S(package, entityName);
            //        GenerateXML_Executable_SQLTask_S2SReAddPK(package, entityName);
            //        GenerateXML_Executable_SQLTask_S2SReAddIndexes(package, entityName);

            //        //Constraints link Executables
            //        int ConstraintNumber = 1;
            //        GenerateXML_PrecedenceConstraint(package, "Truncate Staging Table", "Drop Indexes and PK", ConstraintNumber++);
            //        GenerateXML_PrecedenceConstraint(package, "Drop Indexes and PK", "Data Flow Task_1", ConstraintNumber++);
            //        GenerateXML_PrecedenceConstraint(package, "Data Flow Task_1", "ReAdd PK", ConstraintNumber++);
            //        GenerateXML_PrecedenceConstraint(package, "ReAdd PK", "ReAdd Indexes", ConstraintNumber++);

            //        var result = MessageBox.Show($"This will create {entityName}.dtsx at {packageLocation}\\\n\nWARNING - If that file exists already, it will be OVERWRITTEN!\n\nAre you happy to proceed?", "Warning",
            //                         MessageBoxButtons.YesNo,
            //                         MessageBoxIcon.Question);

            //        if (result == DialogResult.Yes)
            //        {
            //            try
            //            {
            //                SavePackage(package, entityName, packageLocation);
            //                MessageBox.Show("SSIS package created successfully.");
            //            }
            //            catch (Exception e)
            //            {
            //                MessageBox.Show("SSIS package failed to save. Check Source To Staging SSIS Project Location is correct.");
            //            }

            //            addSSISPackageToProject(entityName);
            //        }
            //    }
            //}

            
        }



        public XDocument GetStagingToCRMProject()
        {
            var path = mySettings.StagingToCRMLocationString + "\\StagingToCRM.dtproj";
            XDocument package = null;

            try
            {
                package = XDocument.Load(path);
            }
            catch (Exception e)
            {
                MessageBox.Show("Unable to find a file called \"StagingToCRM.dtproj\" at the Staging To CRM SSIS Project Location. Please ensure you have the correct location set, e.g. C:\\Users\\Chris\\DMSolution\\StagingToCRM\n\n For more information please see the About button.");
            }

            return package;
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