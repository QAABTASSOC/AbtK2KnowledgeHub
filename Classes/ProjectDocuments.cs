﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class ProjectDocuments
    {
        /// <summary>
        /// SharePoint Field : AbtkDocumentId
        /// SharePoint Type : Number
        /// </summary>
        public Int32? DocumentID { get; set; }
        /// <summary>
        /// SharePoint Field : KHProject
        /// SharePoint Type : Lookup
        /// </summary>
        public Int32? ProjectsID { get; set; }
        /// <summary>
        /// SharePoint Field : Title
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// SharePoint Field : KHFormerAbrAuthor
        /// SharePoint Field : KHAbtkAuthor
        /// SharePoint Field : NonAbtAuthors
        /// SharePoint Type : Person and Group
        /// </summary>
        public string Author { get; set; }
        /// <summary>
        /// SharePoint Field : DocumentDate
        /// SharePoint Type : Date and Time
        /// </summary>
        public DateTime? DocumentDate { get; set; }
        public string ProjectNumber { get; set; }
        public string ProjectName { get; set; }
        public string DocumentName { get; set; }


        public string FileSize { get; set; }
    }
}
