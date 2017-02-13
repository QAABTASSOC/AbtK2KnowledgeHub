using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class Staff
    {
        /// <summary>
        /// SharePoint Field : AbtkProjectStaffID
        /// SharePoint Type : Number
        /// </summary>
        public int ProjectStaffID { get; set; }
        /// <summary>
        /// SharePoint Field : KHProject
        /// SharePoint Type : Lookup
        /// </summary>
        public int ProjectsID { get; set; }
        /// <summary>
        /// SharePoint Field : KH_Employee
        /// SharePoint Old Data Field: KHEmployeeName
        /// SharePoint Type : Person or Group
        /// </summary>
        public string  Employee { get; set; }
        /// <summary>
        /// SharePoint Field : BS_Role
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string Role { get; set; }
        /// <summary>
        /// SharePoint Field : StartDate
        /// SharePoint Type : Date Time
        /// </summary>
        public DateTime StartDate { get; set; }
        /// <summary>
        /// SharePoint Field : _EndDate
        /// SharePoint Type : Date Time
        /// </summary>
        public DateTime EndDate { get; set; }
    }
}
