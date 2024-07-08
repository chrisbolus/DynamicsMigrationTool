using Microsoft.SqlServer.Management.Smo;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsMigrationTool
{
    public class EntityAttribute_AdditionalInfo
    {
        public string fieldName;
        public string entityName;
        public string DBDataType;
        public bool stgNotNull = false;
        public string SSISDataType;
        public int? StringLength = null;
        public bool isValidForMigration = true;
        public bool isLookup = false;
        public bool isPrimaryKey = false;
        public bool isDMTField = false;
        public string FieldDerivedFrom;
        public string DynEntityLookupTargets_List;
        public string DynDataType_Readable;

        public EntityAttribute_AdditionalInfo() { }

 

        /// <summary>
        /// Constructor for int
        /// </summary>
        /// <param name="EntityName"></param>
        /// <param name="FieldName"></param>
        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_DMTIntegerField(string EntityName, string FieldName, bool useNotNull = false)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = FieldName;
            EAAI.entityName = EntityName;
            EAAI.DBDataType = "INT";
            EAAI.stgNotNull = useNotNull;
            EAAI.SSISDataType = "i4";
            EAAI.StringLength = null;
            EAAI.isValidForMigration = true;
            EAAI.isLookup = false;
            EAAI.isPrimaryKey = false;
            EAAI.isDMTField = true;

            return EAAI;
        }

        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_DMTSourceField(EntityAttribute_AdditionalInfo EAAI_Existing)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = EAAI_Existing.fieldName + "_Source";
            EAAI.entityName = EAAI_Existing.entityName;
            EAAI.DBDataType = $"NVARCHAR(255)";
            EAAI.stgNotNull = EAAI_Existing.isPrimaryKey ? true : false;
            EAAI.SSISDataType = "wstr";
            EAAI.StringLength = 255;
            EAAI.isValidForMigration = true;
            EAAI.isLookup = true;
            EAAI.isPrimaryKey = EAAI_Existing.isPrimaryKey;
            EAAI.isDMTField = true;
            EAAI.FieldDerivedFrom = EAAI_Existing.fieldName;
            EAAI.DynEntityLookupTargets_List = EAAI_Existing.DynEntityLookupTargets_List;

            return EAAI;
        }
        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_DMTLeaderField(EntityAttribute_AdditionalInfo EAAI_Existing)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = EAAI_Existing.fieldName + "_Leader";
            EAAI.entityName = EAAI_Existing.entityName;
            EAAI.DBDataType = $"NVARCHAR(255)";
            EAAI.SSISDataType = "wstr";
            EAAI.StringLength = 255;
            EAAI.isValidForMigration = true;
            EAAI.isLookup = true;
            EAAI.isPrimaryKey = EAAI_Existing.isPrimaryKey;
            EAAI.isDMTField = true;
            EAAI.FieldDerivedFrom = EAAI_Existing.fieldName;
            EAAI.DynEntityLookupTargets_List = EAAI_Existing.DynEntityLookupTargets_List;

            return EAAI;
        }

        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_StgUniqueIdentifier(string EntityName, string FieldName)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = FieldName;
            EAAI.entityName = EntityName;
            EAAI.DBDataType = "UNIQUEIDENTIFIER";
            EAAI.SSISDataType = "guid";

            return EAAI;
        }

        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_StgDateTime(string EntityName, string FieldName)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = FieldName;
            EAAI.entityName = EntityName;
            EAAI.DBDataType = "DATETIME";
            EAAI.SSISDataType = "dbTimeStamp";

            return EAAI;
        }
        public static EntityAttribute_AdditionalInfo EntityAttribute_AdditionalInfo_StgBoolean(string EntityName, string FieldName)
        {
            var EAAI = new EntityAttribute_AdditionalInfo();

            EAAI.fieldName = FieldName;
            EAAI.entityName = EntityName;
            EAAI.DBDataType = "BIT NOT NULL";
            EAAI.SSISDataType = "bool";

            return EAAI;
        }
    }
}
