using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class Projects
    {
        /// <summary>
        /// SharePoint Field : AbtkProjectId
        /// SharePoint Type : Number
        /// </summary>
        public int? ProjectsID { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectContractNumber
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ContractNumber { get; set; }
        /// <summary>
        /// SharePoint Field : StartDate
        /// SharePoint Type : Date and Time
        /// </summary>
        public DateTime? BeginDate { get; set; }
        /// <summary>
        /// SharePoint Field : _EndDate
        /// SharePoint Type : Date and Time
        /// </summary>
        public DateTime? EndDate { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectOriginalEndDate
        /// SharePoint Type : Date and Time
        /// </summary>
        public DateTime OriginalEndDate { get; set; }
        /// <summary>
        /// SharePoint Field : Title
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ProjectTitle { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectOracleNumber
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ProjectNumber { get; set; }
        /// <summary>
        /// SharePoint Field : BS_Project
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectComments
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string ProjectComments { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectPotentialWorth
        /// SharePoint Type : Currency
        /// </summary>
        public decimal? PotentialWorth { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectContractValue
        /// SharePoint Type : Currency
        /// </summary>
        public decimal? ContractValue { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectCurrentFunding
        /// SharePoint Type : Currency
        /// </summary>
        public decimal CurrentFunding { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectAdditionalReference
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string AdditionalReference { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectIsGoodReference
        /// SharePoint Type : Yes/No
        /// </summary>
        public bool? IsGoodReference { get; set; }

        /// <summary>
        /// SharePoint Field: BS_ProjectDirector
        /// SharePoint Old Data Field: ProjectDirectorOld
        /// Data: Email Id will come as data
        /// SharePoint Type : Person or Group
        /// </summary>
        public string ProjectDirector { get; set; }

        /// <summary>
        /// SharePoint Field: ProjectTechnicalOfficer
        /// SharePoint Old Data Field: ProjectTechnicalOfficerOld
        /// Data: Email Id will come as data
        /// SharePoint Type : Person or Group
        /// </summary>        
        public string TechnicalOfficer { get; set; }
        /// <summary>
        /// SharePoint: ProjectStatus
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ProjectStatus { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectType
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string ProjectType { get; set; }
        /// <summary>
        /// SharePoint Field : Is_x0020_Active
        /// SharePoint Type : Choice
        /// </summary>
        public bool? IsActive { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectIsPrime
        /// SharePoint Type : Yes/No
        /// </summary>
        public bool IsPrime { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectClient
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string Client { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectUltimateClient
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string UltimateClient { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectAgreementId
        /// SharePoint Type : Number
        /// </summary>
        public int? AgreementID { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectAgreementName
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string AgreementName { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectAgreementType
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string AgreementType { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectDivision
        /// SharePoint Type : Choice
        /// </summary>
        public string Division { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectPractice
        /// SharePoint Type : Choice
        /// </summary>
        public string Practice { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectInstClient
        /// SharePoint Type : Multiple Line of Text
        /// </summary>
        public string InstClient { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectFederalAgency
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string FederalAgency { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectAgreementTrackNumber
        /// SharePoint Type : Number
        /// </summary>
        public decimal? AgreementTrackNumber { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectMVTitle
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string MVTitle { get; set; }
        /// <summary>
        /// SharePoint Field : ProjectMMG
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public string MMG { get; set; }
        /// <summary>
        /// SharePoint Field : ParentProject
        /// SharePoint Type : Lookup (self)
        /// </summary>
        public int ParentProject { get; set; }
        /// <summary>
        /// SharePoint Field : Proposal
        /// SharePoint Type : Single Line of Text
        /// </summary>
        public int ProposalID { get; set; }
    }
}
