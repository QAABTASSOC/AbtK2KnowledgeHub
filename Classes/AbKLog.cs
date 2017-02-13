using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class AbKLog
    {
        public int SharePointID { get; set; }
        public long EntityID { get; set; }
        /// <summary>
        /// Category of log (whether task, contact, email, configuration, etc.)
        /// </summary>
        public Enums.MigrationModule Category { get; set; }
        /// <summary>
        /// Type of log Info, Error, etc.
        /// </summary>
        public Enums.LogType Type { get; set; }
        /// <summary>
        /// Operation of log add, update, etc.
        /// </summary>
        public Enums.OperationType OperationType { get; set; }
        /// <summary>
        /// Actual log message
        /// </summary>
        public string Message { get; set; }  
        /// <summary>
        /// 
        /// </summary>
        public Guid CorelationId { get; set; }
        /// <summary>
        /// Created On date
        /// </summary>
        public DateTime CreatedOn { get; set; }   
        public string EntityName { get; set; }    
    }
}
