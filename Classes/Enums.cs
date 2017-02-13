using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public static class Enums
    {

        public enum MigrationModule
        {
            Project = 1,
            ENRTasks = 2,
            Description = 3,
            Staff = 4,
            Documents = 5,
            Tags = 6,
            DatabaseConnection = 7,          
            Metadata = 8,
            Other = 9,
            ENRStaff= 10
        }
        public enum LogType
        {
            Info = 1,
            Error = 2,
            Warning = 3
        }
        public enum OperationType
        {
            NotKnown = 1,
            Add = 2,
            Update = 3,
            ApplyingTags = 4,
            DuplicatingStaffOnENR = 5,
            ApplyingMetadata = 6

        }

    }
}
