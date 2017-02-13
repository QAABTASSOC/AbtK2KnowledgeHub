using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class ProjectDescription
    {
        /// <summary>
        /// SharePoint Field : AbtKDescriptionId
        /// SharePoint Type : Number
        /// </summary>
        public int DescriptionID { get; set; }
        /// <summary>
        /// SharePoint Field : KHProject
        /// SharePoint Type : Lookup
        /// </summary>
        public int ProjectsID { get; set; }
        /// <summary>
        /// SharePoint Field : DescriptionType
        /// SharePoint Type : Choice
        /// </summary>
        public int DescriptionType { get; set; }
        /// <summary>
        /// SharePoint Field : Title
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_Overview
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string Overview { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_ScopeOfWork
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string ScopeOfWork { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_KeyDeliverable
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string KeyDeliverables { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_InnovativeToolsorResources
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string Innovative { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_Accomplishments
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string Accomplishments { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_Problems
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string Problems { get; set; }
        /// <summary>
        /// SharePoint Field : PDS_Awards
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string Awards { get; set; }
        /// <summary>
        /// SharePoint Field : Is_x0020_Active
        /// SharePoint Type : bool
        /// </summary>
        public bool IsActive { get; set; }
    }
}
