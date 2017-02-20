﻿using AbtK2KnowledgeHub_OneTime.Classes;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;

namespace AbtK2KnowledgeHub_OneTime
{
    class Program
    {
        Guid correlationId = Guid.NewGuid();
        string knowledgeHubWebUrl = Helper.GetAppSettingValue(Constants.KnowledgeHubSiteUrlKey);
        string emailId = Helper.GetAppSettingValue(Constants.KnowledgeHubEmailId);
        string password = Helper.GetAppSettingValue(Constants.KnowledgeHubPassword);
      
        private static bool isDoneExecuting = false;

        //path for the logs to be written
        private static string logPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                 + @"\KH_MIGRATION\";

        //contains all of the valid formats for the test
        public static HashSet<string> extensions = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase)
        { ".WPD",".wpd",".doc", ".docx", ".xlsx", ".xls",".pdf",".txt", ".pptx",".ppt", ".jpg", ".zip",".png",".rar" };

        public static Dictionary<string, string> ignoredRecords = new Dictionary<string, string>();
        public static Dictionary<string, Projects> ProjectsFromDB = new Dictionary<string, Projects>();
        public static Dictionary<Int32, Proposals> ProposalsFromDB = new Dictionary<Int32, Proposals>();


        static void Main(string[] args)
        {
            Program program = new Program();
            try
            {
                if (string.IsNullOrEmpty(program.knowledgeHubWebUrl) || string.IsNullOrEmpty(program.emailId) || string.IsNullOrEmpty(program.password))
                {
                    Program.LogNDisplay("Please check app.config and supply values for weburl, email id and password");
                    return;
                }
                //Projects
               // program.ReadProjectsFromSQL();
               // program.ReadProjectDescriptionFromSQL();
               // program.ReadProjectDocumentsFromSQL();

                //sharepoint extract
               // ExcelReader.ReadConfig("Projects", "Projects");
               // ExcelReader.ReadConfig("Descriptions", "Projects");
                ExcelReader.ReadConfig("Documents", "Projects");

                //Proposals
                program.ReadProposalsFromSQL();

                //sharedpoint extract
                ExcelReader.ReadConfig("Proposals", "Proposals and Repcap extract");


                Console.WriteLine("Import is complete. Press any key to exit.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Program.LogNDisplay("Uncatched Exception: " + ex.Message);
            }

        }
       
        public void ReadProjectsFromSQL()
        {
            int countr = 0;
            Console.WriteLine("Begin Reading Projects from the SQL \n");

            using (SqlConnection sqlConnetion = new SqlConnection("Data Source = 10.221.100.52; Initial Catalog = abtknowledge; User ID = abtknowledge; Password = 2SxBZD3er63C; persist security info=True;"))
            {
                string queryStatement =  "SELECT* FROM " + Helper.GetAppSettingValue(Constants.ProjectViewKey) ; 
                using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                {
                    sqlConnetion.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    int count = 0; int i = 0;
                    while (reader.Read())
                    {
                        string projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                        Projects thisProject = new Projects();

                        //Project Number
                        thisProject.ProjectNumber = projectNumber;
                        // is Active
                        thisProject.IsActive = Helper.SafeGetBool(reader, "IsActive");
                        //Project Name
                        thisProject.ProjectName = Helper.SafeGetString(reader, "ProjectName");
                        //project ID
                        thisProject.ProjectsID = Helper.SafeGetInt32(reader, "ProjectsID");
                        //Project Title
                        thisProject.ProjectTitle = Helper.SafeGetString(reader, "ProjectTitle");
                        //Is this project a good reference?
                        thisProject.IsGoodReference = Helper.SafeGetBool(reader, "IsGoodReference");
                        //Is Abt the Prime? Yes
                        thisProject.IsPrime = Convert.ToBoolean(Helper.SafeGetBool(reader, "IsPrime")) ? true : false;
                        //Client 
                        thisProject.Client = Helper.SafeGetString(reader, "Client");
                        //Ultimate Client
                        thisProject.UltimateClient = Helper.SafeGetString(reader, "UltimateClient");
                        //Contract Number
                        thisProject.ContractNumber = Helper.SafeGetString(reader, "ContractNumber");
                        //Additional Reference
                        thisProject.AdditionalReference = Helper.SafeGetString(reader, "AdditionalReference");
                        //Agreement Name 
                        thisProject.AgreementName = Helper.SafeGetString(reader, "AgreementName");
                        //Agreement Type 
                        thisProject.AgreementType = Helper.SafeGetString(reader, "AgreementType");
                        //Agreement Id
                        thisProject.AgreementID = Helper.SafeGetInt32(reader, "AgreementID");
                        //Begin Date 
                        thisProject.BeginDate = Helper.SafeGetDateTime(reader, "BeginDate");
                        //End Date 
                        thisProject.EndDate = Helper.SafeGetDateTime(reader, "EndDate");
                        //Potential Worth
                        thisProject.PotentialWorth = Helper.SafeGetDecimal(reader, "PotentialWorth");
                        //Award Amount
                       // thisProject.AwardAmount = Helper.SafeGetDecimal(reader, "AwardAmount");
                        //Funded Amount
                        //Division  
                        thisProject.Division = GetDivision(Helper.SafeGetString(reader, "Division"));
                        //Practice  
                        thisProject.Practice = GetPractice(Helper.SafeGetString(reader, "Practice"));
                        //Project Director 
                        thisProject.ProjectDirector = Helper.SafeGetString(reader, "ProjectDirector");
                        //Project Director Name  
                        //Technical Officer
                        thisProject.TechnicalOfficer = Helper.SafeGetString(reader, "TechnicalOfficer");
                        //Technical Officer Name
                        //Parent Project 
                        string parentProjectNumber = Helper.SafeGetString(reader, "ParentProjectNumber");
                        //Is Active ? (Y / N)  Yes
                        thisProject.IsPrimeText = (bool)thisProject.IsActive ? "Yes" : "No";
                        thisProject.InstClient = Helper.SafeGetString(reader, "InstClient");
                        thisProject.FederalAgency = Helper.SafeGetString(reader, "FederalAgency");
                        thisProject.AgreementTrackNumber = Helper.SafeGetDecimal(reader, "AgreementTrackNumber");
                        thisProject.MVTitle = Helper.SafeGetString(reader, "MVTitle");
                        thisProject.MMG = Helper.SafeGetString(reader, "MMG");

                        string ProposalOracleNumber = Helper.SafeGetString(reader, "Proposalnumber");
                        string isGoodRef;
                        if (thisProject.IsGoodReference.HasValue)
                            isGoodRef = (bool)thisProject.IsGoodReference ? "Yes" : "No";
                        else
                            isGoodRef = "Unknown";

                        thisProject.ProjectComments = Helper.SafeGetString(reader, "ProjectComments");
                        thisProject.ContractValue = Helper.SafeGetDecimal(reader, "ContractValue");

                        //add to index map
                        if (!ProjectsFromDB.ContainsKey(projectNumber))
                        {
                            ProjectsFromDB.Add(projectNumber, thisProject);
                            Program.LogNDisplay(projectNumber + " have been added to the Dictionary #" + count);
                        }
                        else
                        {
                            Program.LogNDisplay("the file: " + projectNumber + " have been previously processed: " + " \n index: " + count);
                        }
                        count++;
                    }
                    sqlConnetion.Close();
                }
            }
        }
        public void ReadProjectDescriptionFromSQL()
        {

            Program.LogNDisplay("\n Begin Reading Descriptions from the SQL \n");
            try
            {

                using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                {
                    string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectDescriptionViewKey);

                    using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                    {
                        sqlConnetion.Open();
                        SqlDataReader reader = command.ExecuteReader();
                        int count = 0;

                        while (reader.Read())
                        {
                            try
                            {
                                ProjectDescription thisDescription = new ProjectDescription();

                                string projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                                thisDescription.ProjectNumber = projectNumber;
                                thisDescription.Title = Helper.SafeGetString(reader, "Title");
                                //unique id field in the view in sql
                                thisDescription.DescriptionID = Helper.SafeGetInt64(reader, "OverviewID");
                                thisDescription.ProjectsID = Helper.SafeGetInt32(reader, "DescriptionID");
                                thisDescription.DescriptionType = Helper.SafeGetInt32(reader, "DescriptionType");


                                //add to index map
                                if (ProjectsFromDB.ContainsKey(projectNumber))
                                {
                                    if (!ProjectsFromDB[projectNumber].DescriptionContainsKey(Convert.ToString(thisDescription.DescriptionID)))
                                    {

                                        //add Description to Project
                                        ProjectsFromDB[projectNumber].SetDescription(projectNumber, thisDescription);
                                        Program.LogNDisplay("Description ID: " + thisDescription.DescriptionID + " for Project #" +
                                                            projectNumber + " have been added to the Dictionary #" + count);
                                    }
                                    else
                                    {
                                        Program.LogNDisplay("Description ID " + thisDescription.DescriptionID + " for project" +
                                                            projectNumber + " is already there #" + count);
                                    }
                                    count++;
                                }
                            }
                            catch (Exception e)
                            {
                                Program.LogNDisplay("Failed to catch Description from SQL: "+e.Message);
                            }
                        }
                        sqlConnetion.Close();
                    }
                }
            }
            catch (Exception e)
            {
                Program.LogNDisplay("Could not connect to ProjectDescriptionViewKey " + e.Message);
            }
        }
        public void ReadProjectDocumentsFromSQL()
        {
            Console.WriteLine("Applying Metadata to documents");
                           
               using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectDocumentsViewKey);
                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader();
                            int count = 0; int i = 0; string lastProjectNumber = string.Empty;

                            while (reader.Read())
                            {
                              
                                string projectNumber = string.Empty;
                                string fileName = string.Empty;
                                string documentTitle = string.Empty;

                                try
                                {
                                    ProjectDocuments thisDocument = new ProjectDocuments();        
                                    projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                                    fileName = Helper.SafeGetString(reader, "UploadedFileName");
                                    documentTitle = Helper.SafeGetString(reader, "Title");

                                    thisDocument.DocumentID = Helper.SafeGetInt32(reader, "FilesID");
                                    thisDocument.Title  = String.IsNullOrEmpty(documentTitle) ? "" : StringExt.Truncate(documentTitle, 255);
                                    thisDocument.ProjectNumber = projectNumber;
                                    thisDocument.DocumentName = fileName;
                                    thisDocument.Author = Helper.SafeGetString(reader, "Author");
                                    thisDocument.ProjectsID = Helper.SafeGetInt32(reader, "ProjectsID");
                                    thisDocument.DocumentDate = Helper.SafeGetDateTime(reader, "FileDate");

                            //add to index map
                            if (ProjectsFromDB.ContainsKey(projectNumber))
                            {
                                if (!ProjectsFromDB[projectNumber].DocumentContainsKey(Convert.ToString(thisDocument.DocumentName)))
                                {

                                    //add Description to Project
                                    ProjectsFromDB[projectNumber].SetDocuments(projectNumber, thisDocument);
                                    Program.LogNDisplay("Document ID: " + thisDocument.DocumentID+ "_" +thisDocument.DocumentName + " for Project #" +
                                                        projectNumber + " have been added to the Dictionary #" + count);
                                }
                                else
                                {
                                    Program.LogNDisplay("Document ID " + thisDocument.DocumentID + "_" + thisDocument.DocumentName+ " for project" +
                                                        projectNumber + " is already there #" + count);
                                }
                                count++;
                            }

                        }

                        catch (Exception e)
                        {
                            Program.LogNDisplay("Failed to catch  Document from SQL: " + e.Message);
                        }

                    }
                            sqlConnetion.Close();
                        }
                    }
            }

        public void ReadProposalsFromSQL()
        {
            int countr = 0;
            Console.WriteLine("Begin Reading Projects from the SQL \n");
            using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
            {
                string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProposalViewKey);
                using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                {
                    sqlConnetion.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    int count = 0;
                    while (reader.Read())
                    {
                        try
                        {

                            string projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                            Proposals thisProposal = new Proposals();

                            //"ProposalOracleNumber"] = proposalNumber;
                            //"ProposalAbtkId"] = proposalId;
                            thisProposal.ProjectNumber = projectNumber;
                            thisProposal.ProposalNumber = Helper.SafeGetString(reader, "ProposalNumber");
                            thisProposal.ProposalsID = Helper.SafeGetInt32(reader, "ProposalsID");

                            thisProposal.ProposalTitle = Helper.SafeGetString(reader, "FullTitle");
                            thisProposal.ProposalTitle = string.IsNullOrEmpty(thisProposal.ProposalTitle) ? "" : StringExt.Truncate(thisProposal.ProposalTitle, 255);
                            thisProposal.RPFTitle = Helper.SafeGetString(reader, "RFPTitle");
                            thisProposal.IsActive = Helper.SafeGetBool(reader, "IsActive");
                            thisProposal.IsGoodExample = Helper.SafeGetBool(reader, "IsGoodExample");

                            thisProposal.ProposalManager = Helper.SafeGetString(reader, "ProposalManager");
                            string proposalLead = Helper.SafeGetString(reader, "Lead");

                            thisProposal.DueDate = Helper.SafeGetDateTime(reader, "DueDate");
                            thisProposal.ProposalName = Helper.SafeGetString(reader, "ProposalName");
                            thisProposal.RPFTitle = string.IsNullOrEmpty(thisProposal.RPFTitle) ? "" : StringExt.Truncate(thisProposal.RPFTitle, 255);
                            thisProposal.RPFNumber = Helper.SafeGetString(reader, "RFPNumber");
                            thisProposal.ProposalComments = Helper.SafeGetString(reader, "ProposalComment");
                            thisProposal.Summary = Helper.SafeGetString(reader, "ProposalSummary");
                            thisProposal.ProposalWorth = Helper.SafeGetDecimal(reader, "ProposalWorth");
                            thisProposal.ProposalHasWon = Convert.ToBoolean(Helper.SafeGetBool(reader, "HasWon")) ? true : false;
                            thisProposal.Client = Helper.SafeGetString(reader, "Client");


                            thisProposal.AgreementTrackNumber = Helper.SafeGetDecimal(reader, "AgreementTrackNumber");
                            //    thisProject.ProjectComments = Helper.SafeGetString(reader, "ProjectComments");
                            //  thisProject.ContractValue = Helper.SafeGetDecimal(reader, "ContractValue");

                            thisProposal.UltimateClient = Helper.SafeGetString(reader, "UltimateClient");
                            thisProposal.AgreementID = Helper.SafeGetInt64(reader, "AgreementID");
                            thisProposal.AgreementName = Helper.SafeGetString(reader, "AgreementName");
                            thisProposal.AgreementType = Helper.SafeGetString(reader, "AgreementType");
                            thisProposal.Division = GetDivision(Helper.SafeGetString(reader, "Division"));
                            thisProposal.Practice = GetPractice(Helper.SafeGetString(reader, "Practice"));
                            thisProposal.FederalAgency = Helper.SafeGetString(reader, "FederalAgency");
                            thisProposal.AgreementTrackNumber = Helper.SafeGetInt64(reader, "AgreementTrackNumber");
                            thisProposal.MMG = Helper.SafeGetString(reader, "MMG");
                            thisProposal.ProposalWinStatus = Helper.SafeGetString(reader, "WinStatus");
                            thisProposal.NoDocumentSubmitteds = Convert.ToBoolean(Helper.SafeGetBool(reader, "NoDocumentSubmitted")) ? true : false;
                            thisProposal.IsPrime = Convert.ToBoolean(Helper.SafeGetBool(reader, "IsPrime")) ? true : false;

                            //Is Active ? (Y / N)  Yes
                            thisProposal.IsPrimeText = (bool)thisProposal.IsActive ? "Yes" : "No";

                            if (thisProposal.IsActive.HasValue)
                                thisProposal.IsActiveText = thisProposal.IsActive.HasValue ? "Yes" : "No";
                            else
                                thisProposal.IsActiveText = "No";

                            if (thisProposal.IsGoodExample.HasValue)
                                thisProposal.IsGoodExampleText = thisProposal.IsGoodExample.Value ? "Yes" : "No";
                            else
                                thisProposal.IsGoodExampleText = "Not Known";

                            //add to index map
                            if (!ProposalsFromDB.ContainsKey((int)thisProposal.ProposalsID))
                            {
                                 ProposalsFromDB.Add((int)thisProposal.ProposalsID, thisProposal);
                                Program.LogNDisplay("Proposal: "+(int)thisProposal.ProposalsID + " for the project" + projectNumber + " have been added to the Dictionary #" + count);
                            }
                            else
                            {
                                Program.LogNDisplay("the file: " + (int)thisProposal.ProposalsID + " have been previously processed #" + count);
                            }
                        }
                        catch(Exception e)
                        {
                            Program.LogNDisplay("Error While reading from SQL" + count +"\n " +e.Message);
                        }
                        count++;
                        
                    }
                    sqlConnetion.Close();
                }
            }
        }


        public static void LogNDisplay(string action, long elapsedTipe)
        {
            using (StreamWriter w = System.IO.File.AppendText(logPath + "ABTKMigration.txt"))
            {
                Log(action + ": " + TimeSpan.FromMilliseconds(elapsedTipe).ToString(), w);
            }
            using (StreamReader r = System.IO.File.OpenText(logPath + "ABTKMigration.txt"))
            {
                DumpLog(r);
            }
        }
        /// <summary>
        ///  log and Display
        /// </summary>
        /// <param name="action">action to be logged i.e (Move, Delete, Crete, etc..)</param>
        public static void LogNDisplay(string action)
        {
            using (StreamWriter w = System.IO.File.AppendText(logPath + "ABTKMigration.txt"))
            {
                Log(action, w);
            }
            //using (StreamReader r = File.OpenText(logPath + "KH_FILES_AND_PATHS_LOG.txt"))
            //{
            //    DumpLog(r);
            //}
        }
        public static void CleanLogNDisplay(string action)
        {
            using (StreamWriter w = System.IO.File.AppendText(logPath + "ABTKMigration.txt"))
            {
                CleanLog(action, w);
            }
            //using (StreamReader r = File.OpenText(logPath + "KH_FILES_AND_PATHS_LOG.txt"))
            //{
            //    DumpLog(r);
            //}
        }
        /// <summary>
        /// inserts log messages
        /// </summary>
        /// <param name="logMessage"></param>
        /// <param name="w"></param>
        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("\r\nEntry at : ");
            w.WriteLine("{0}", DateTime.Now.ToString());
            // w.WriteLine("  ");
            w.WriteLine("{0}", logMessage);
            Console.WriteLine(logMessage);
            //  Console.Clear();
        }
        public static void CleanLog(string logMessage, TextWriter w)
        {
            w.WriteLine("{0}\n", logMessage);
            Console.WriteLine(logMessage);
            //  Console.Clear();
        }
        /// <summary>
        /// display the log to standar output
        /// </summary>
        /// <param name="r"></param>
        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }

        //DEEPA'S Code Bellow
       
        public string GetDivision(string division)
        {
            switch (division)
            {
                case "International Health":
                    return "International Health Division (IHD)";
                case "US Health":
                    return "US Health(USH)";
                case "Social & Economic Policy":
                    return "Social & Economic Policy Division(SEP)";
                case "International Economic Growth":
                    return "International Economic Growth Division(IEG)";
                case "Environment & Resources":
                    return "Environmental & Natural Resources Division(ENR)";
                default:
                    return division;
            }

        }

        public string GetPractice(string practice)
        {
            switch (practice)
            {
                case "Education Evaluation Prc":
                    return "Education(EDU)";
                case "Income Security and Workforce Prc":
                    return "Income Security & Workforce(ISW)";
                case "Housing Prc":
                    return "Housing(HPP)";
                case "Health Policy Prc":
                    return "Health Policy (HP)";
                case "Public Health & Epidemiology Prc":
                    return "Public Health & Epidemiology(PHE)";
                case "Behavioral Health Prc":
                    return "Behavioral Health(BH)";
                default:
                    return practice;
            }
        }
    }
}
