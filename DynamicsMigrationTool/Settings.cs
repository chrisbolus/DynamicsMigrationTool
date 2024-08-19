using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsMigrationTool
{
    /// <summary>
    /// This class can help you to store settings for your plugin
    /// </summary>
    /// <remarks>
    /// This class must be XML serializable
    /// </remarks>
    public class Settings
    {
        public string LastUsedOrganizationWebappUrl { get; set; }
        public string StagingDBConnectionString { get; set; }
        public string StagingDBSchema { get; set; }
        public string ImportSchema { get; set; }
        public string SourceDBConnectionString { get; set; }
        public string SourceDBSchema { get; set; }
        public string SourceToStagingLocationString { get; set; }
        public string CRMToStagingLocationString { get; set; }
        public bool UseDataTransforms { get; set; }
    }
}