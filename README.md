# DynamicsMigrationTool

This tool is designed to aid migration into D365 using a Dynamics connection and a local MSSQL databases. The tool uses the Dynamics Entity metadata in order to create various components which can be used in conjuction to load data into Dynamics.

Current Functionality:
-Generation of Template Source views in Source MSSQL Database, based on the metadata of an entity.
-Generation of Staging tables in Staging MSSQL Database, based on the metadata of an entity.
-Generation of Source to Staging SSIS packages, using Source DB view as the source and the Staging Table as the target.

Future Functionality:
-Generation of Staging to CRM SSIS packages (using the KingswaySoft adaptor). using Staging table as the source and the Dynamics entity as the target.