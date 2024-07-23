using Microsoft.SqlServer.Management.Smo;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Activities.Expressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace DynamicsMigrationTool
{
    public static class CRMHelper
    {
        public static EntityAttribute_AdditionalInfo Get_EntityAttribute_AdditionalInfo(IOrganizationService organizationService, AttributeMetadata attribute)
        {
            var EntAAI = new EntityAttribute_AdditionalInfo();
            EntAAI.fieldName = attribute.LogicalName;
            EntAAI.entityName = attribute.EntityLogicalName;
            EntAAI.DynDataType_Readable = attribute.AttributeType?.ToString();      //this is modified in a few of cases below.

            if ((attribute.AttributeOf != null && attribute.AttributeType == AttributeTypeCode.String)
                || (attribute.IsValidForCreate == false && attribute.IsValidForUpdate == false)
                || (attribute.IsLogical == true && attribute.IsPrimaryId == true)
                || attribute.AttributeType == AttributeTypeCode.PartyList
                || attribute.AttributeType == AttributeTypeCode.ManagedProperty)
            {
                //Bad Records
                EntAAI.isValidForMigration = false;
            }
            else if (attribute.AttributeType == AttributeTypeCode.BigInt)
            {
                EntAAI.DBDataType = "BIGINT";
                EntAAI.SSISDataType = "i8";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Boolean
                || attribute.AttributeType == AttributeTypeCode.EntityName
                || attribute.AttributeType == AttributeTypeCode.Picklist
                || attribute.AttributeType == AttributeTypeCode.State
                || attribute.AttributeType == AttributeTypeCode.Status)
            {
                EntAAI.DBDataType = "NVARCHAR(255)";
                EntAAI.SSISDataType = "wstr";
                EntAAI.StringLength = 255;
            }
            else if (attribute.AttributeType == AttributeTypeCode.Customer
                || attribute.AttributeType == AttributeTypeCode.Lookup
                || attribute.AttributeType == AttributeTypeCode.Owner
                || attribute.AttributeType == AttributeTypeCode.Uniqueidentifier)
            {
                EntAAI.DBDataType = "UNIQUEIDENTIFIER";
                EntAAI.SSISDataType = "guid";
                EntAAI.isLookup = true;
                EntAAI.DynEntityLookupTargets_Count = GetEntityLookupTargetCount(attribute);
                EntAAI.DynEntityLookupTargets_List = GetEntityLookupTargetsList(attribute);

                if (attribute.AttributeType == AttributeTypeCode.Uniqueidentifier)
                {
                    var entityMetadata = organizationService.GetEntityMetadata(attribute.EntityLogicalName);
                    if (entityMetadata.PrimaryIdAttribute == attribute.LogicalName)
                    {
                        EntAAI.isPrimaryKey = true;
                    }
                }
            }
            else if (attribute.AttributeType == AttributeTypeCode.DateTime)
            {
                EntAAI.DBDataType = "DATETIME";
                EntAAI.SSISDataType = "dbTimeStamp";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Decimal)
            {
                EntAAI.DBDataType = "DECIMAL(23, 10)";
                EntAAI.SSISDataType = "numeric";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Double)
            {
                EntAAI.DBDataType = "FLOAT";
                EntAAI.SSISDataType = "r8";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Integer)
            {
                EntAAI.DBDataType = "INT";
                EntAAI.SSISDataType = "i4";
            }
            else if (attribute.AttributeType == AttributeTypeCode.Memo)
            {
                EntAAI.DBDataType = "NVARCHAR(MAX)";
                EntAAI.SSISDataType = "nText";
                EntAAI.StringLength = -1;
            }
            else if (attribute.AttributeType == AttributeTypeCode.Money)
            {
                EntAAI.DBDataType = "MONEY";
                EntAAI.SSISDataType = "cy";
            }
            else if (attribute.AttributeType == AttributeTypeCode.String)
            {
                var strlength = (int)attribute.GetType().GetProperty("MaxLength").GetValue(attribute);
                EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + $"({strlength})";
                if (strlength > 4000)
                {
                    EntAAI.DBDataType = "NVARCHAR(MAX)";
                    EntAAI.SSISDataType = "nText";
                    EntAAI.StringLength = -1;
                }
                else
                {
                    EntAAI.DBDataType = $"NVARCHAR({strlength})";
                    EntAAI.SSISDataType = "wstr";
                    EntAAI.StringLength = strlength;
                }
            }
            else if (attribute.AttributeType == AttributeTypeCode.Virtual)
            {
                if (attribute.AttributeTypeName.Value == "MultiSelectPicklistType")
                {
                    EntAAI.DBDataType = "NVARCHAR(1000)";
                    EntAAI.SSISDataType = "wstr";
                    EntAAI.StringLength = 1000;
                    EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + " - " + attribute.AttributeTypeName.Value;
                }
                else if (attribute.AttributeTypeName.Value == "ImageType")
                {
                    EntAAI.DBDataType = "IMAGE";
                    EntAAI.SSISDataType = "image";
                    EntAAI.DynDataType_Readable = EntAAI.DynDataType_Readable + " - " + attribute.AttributeTypeName.Value;
                }
                else
                {
                    //attribute.AttributeTypeName.Value == "VirtualType"
                    EntAAI.isValidForMigration = false;
                }
            }
            else
            {
                //In theory this should never be hit.
                EntAAI.isValidForMigration = false;
            }

            return EntAAI;

        }


        public static string GetEntityLookupTargetsList(AttributeMetadata attribute)
        {
            if (attribute.GetType().GetProperty("Targets") != null)
            {
                var targets = (string[])attribute.GetType().GetProperty("Targets").GetValue(attribute);

                return String.Join(",", targets);
            }
            else return null;
        }
        public static int? GetEntityLookupTargetCount(AttributeMetadata attribute)
        {
            if (attribute.GetType().GetProperty("Targets") != null)
            {
                var targets = (string[])attribute.GetType().GetProperty("Targets").GetValue(attribute);

                return targets.Count();
            }
            else return null;
        }

        public static List<EntityAttribute_AdditionalInfo> GetFullFieldList(IOrganizationService Service, EntityMetadata entity, bool includeStagingFields = false)
        {
            var eList = new List<EntityAttribute_AdditionalInfo>();

            foreach(var attribute in entity.Attributes.OrderBy(a => a.LogicalName))
            {
                EntityAttribute_AdditionalInfo entAAI = CRMHelper.Get_EntityAttribute_AdditionalInfo(Service, attribute);

                if (entAAI.isValidForMigration) 
                {

                    if (entAAI.isLookup)
                    {

                        eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTSourceField(entAAI));
                        if (entAAI.isPrimaryKey)
                        {
                            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTIntegerField(entAAI.entityName, "Leader_System_Id"));
                            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTLeaderField(entAAI));
                        }
                        else
                        {
                            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTLookupTypeField(entAAI));
                            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTDynamicsLookupFields(entAAI));
                        }
                    }
                    else
                    {
                        eList.Add(entAAI);
                    }
                }
            }

            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTIntegerField(entity.LogicalName, "Source_System_Id", true));
            eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_DMTIntegerField(entity.LogicalName, "Processing_Status", true));

            if (includeStagingFields)
            {
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgBoolean(entity.LogicalName, "FlagCreate"));
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgBoolean(entity.LogicalName, "FlagUpdate"));
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgBoolean(entity.LogicalName, "FlagDelete"));
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgUniqueIdentifier(entity.LogicalName, "DynCreateId"));
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgUniqueIdentifier(entity.LogicalName, "DynUpdateId"));
                //eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgUniqueIdentifier(entity.LogicalName, "DynDeleteId"));
                eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgUniqueIdentifier(entity.LogicalName, "DynId"));
                eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgDateTime(entity.LogicalName, "DateCreate"));
                eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgDateTime(entity.LogicalName, "DateUpdate"));
                eList.Add(EntityAttribute_AdditionalInfo.EntityAttribute_AdditionalInfo_StgDateTime(entity.LogicalName, "DateDelete"));
            }

            return eList;


        }



    }
}
