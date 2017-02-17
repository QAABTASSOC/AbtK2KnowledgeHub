using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    class Proposals
    {
        public string ProjectNumber{ get; set; }
        public int? ProposalsID { get; set; }
        public string ProposalName { get; set; }
        public string ProposalTitle { get; set; }
        public DateTime? DueDate { get; set; }
        public DateTime? LastUpDate { get; set; }
        public DateTime? StaffEditDate { get; set; }
        public string ProposalNumber { get; set; }
        public string ProposalManager { get; set; }
        public string ProposalComments { get; set; }
        public string Comments { get; set; }
        public string Summary { get; set; }
        public string Client { get; set; }
        public bool? IsPrime { get; set; }
        public bool? IsActive { get; set; }
        public string IsActiveText { get; set; }
        public string IsPrimeText { get; set; }
        public string UltimateClient{ get; set; }
        public bool? IsGoodExample { get; set; }
        public string IsGoodExampleText { get; set; }    
        public string ContractType{ get; set; }
        public string RFPNumber { get; set; }
        public string RPFTitle { get; set; }
        public string RPFNumber { get; set; }
        public Decimal? ProposalWorth { get; set; }
        public bool? ProposalHasWon { get; set; }
        public string ProposalWinStatus { get; set; }
        public bool? NoDocumentSubmitteds { get; set; }
        public Int64? AgreementID { get; set; }
        public string AgreementName { get; set; }
        public string AgreementType { get; set; }
        public string Division { get; set; }
        public string Practice { get; set; }
        public string InstClient { get; set; }
        public string FederalAgency { get; set; }
        public decimal? AgreementTrackNumber { get; set; }
        public string MVTitle { get; set; }
        public string MMG { get; set; }
        public Int32? OracleProjectNumber { get; set; }
        public int? ProposalID { get; set; }
        public Int32? OracleProposalNumber { get; set; }
    
    }
}
