using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    class ExcelReader
    {
        private static Excel.Application excelApp;
        private static Excel.Workbook excelWorkbook;
        private static Excel.Worksheet excelWorksheet;

        public static Dictionary<string, Projects> ExcelProjectsDictionary = new Dictionary<string, Projects>();

        public static int ReadConfig(String type)
        {
            excelApp = new Excel.Application();

            try
            {   //TO USE EXCEL: PUT THE PATH IN THE LINE BELOW
                switch (type)
                {
                    case "Proposals":
                        excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\Proposals.xlsx", true, true);
                        excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Proposals");
                        break;

                    case "ProposalsDocuments":
                        excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\Proposals.xlsx", true, true);
                        excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("ProposalsDocuments");
                        break;
                    case "Projects":
                        excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\Projects.xlsx", true, true);
                        excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Projects");
                        break;
                    case "Descriptions":
                        excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\Projects.xlsx", true, true);
                        excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("ProjectDescriptions");
                        break;
                    case "Documents":
                        excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\Projects.xlsx", true, true);
                        excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("ProjectDocuments");
                        break;

                }

            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
            }

            switch (type)
            {
                case "Projects":
                    LoadSharePointPojectsExtract();
                    Program.LogNDisplay("\n Finished Loading Sharepoint Projects \n");
                    break;
                case "Descriptions":
                    LoadSharePointProjectDescriptionsExtract();
                    Program.LogNDisplay("\n Finished Loading Sharepoint Projects Descriptions \n");
                    break;

                case "Documents":
                    LoadSharePointPojectsDocumentsExtract();
                    Program.LogNDisplay("\n Finished Loading Sharepoint Projects Documents \n");
                    break;

                case "Proposals":
                    LoadSharePointProposalsExtract();
                     Program.LogNDisplay("\n Finished Loading Sharepoint Proposals \n");
                    break;

                case "ProposalsDocuments":
                    LoadSharePointProposalDocumentsExtract();
                    Program.LogNDisplay("\n Finished Loading Sharepoint Proposals \n");
                    break;
            }

            excelWorkbook.Close();
            excelApp.Quit();

            //release from memory
            ReleaseObject(excelWorksheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            return 0;
        }
        public static void LoadSharePointPojectsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Projects \n");

            while ((string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value != null)
            {
                Projects project = new Projects();
                string projectNumber = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                try
                {
                    //load row into memory
                    project.ProjectNumber = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                    project.ProjectTitle = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                    project.ProjectName = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                   // project.Division = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    project.Client = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    project.UltimateClient = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                    project.IsPrime = (bool?)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                    project.BeginDate = (DateTime?)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                    project.EndDate = (DateTime?)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                    project.OriginalEndDate = (DateTime?)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                    project.ContractNumber = (string)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                    project.AdditionalReference = (string)(excelWorksheet.Cells[row, 11] as Excel.Range).Value;
                    project.PotentialWorth = (Decimal?)(excelWorksheet.Cells[row, 12] as Excel.Range).Value;
                    project.AwardAmount = (Decimal?)(excelWorksheet.Cells[row, 13] as Excel.Range).Value;
                    project.ProjectType = (string)(excelWorksheet.Cells[row, 14] as Excel.Range).Value;
                    project.FundedAmount = (Decimal?)(excelWorksheet.Cells[row, 15] as Excel.Range).Value;
                    project.ProjectDirector = (string)(excelWorksheet.Cells[row, 16] as Excel.Range).Value;
                    project.ProjectDirectorName = (string)(excelWorksheet.Cells[row, 17] as Excel.Range).Value;
                    project.TechnicalOfficer = (string)(excelWorksheet.Cells[row, 18] as Excel.Range).Value;
                    project.TechnicalOfficerName = (string)(excelWorksheet.Cells[row, 19] as Excel.Range).Value;
                  //  project.Practice = (string)(excelWorksheet.Cells[row, 20] as Excel.Range).Value;
                    project.ProjectStatus = (string)(excelWorksheet.Cells[row, 20] as Excel.Range).Value;
                    project.IsGoodReferenceText = (string)(excelWorksheet.Cells[row, 21] as Excel.Range).Value;
                    project.ParentProject = (string)(excelWorksheet.Cells[row, 22] as Excel.Range).Value;

                    if (((string)(excelWorksheet.Cells[row, 23] as Excel.Range).Value).ToUpper().Equals("YES"))
                        project.IsActive = true;
                    else
                        project.IsActive = false;

                    project.ProjectComments = (string)(excelWorksheet.Cells[row, 24] as Excel.Range).Value;
                    project.AgreementID = (int?)(excelWorksheet.Cells[row, 25] as Excel.Range).Value;
                    project.AgreementName = (string)(excelWorksheet.Cells[row, 26] as Excel.Range).Value;
                    project.AgreementTrackNumber = Convert.ToInt32((string)(excelWorksheet.Cells[row, 27] as Excel.Range).Value);
                    project.AgreementType = (string)(excelWorksheet.Cells[row, 28] as Excel.Range).Value;
                    project.FederalAgency = (string)(excelWorksheet.Cells[row, 29] as Excel.Range).Value;
                    project.MMG = (string)(excelWorksheet.Cells[row, 30] as Excel.Range).Value;
                    project.InstClient = (string)(excelWorksheet.Cells[row, 31] as Excel.Range).Value;
                    project.MVTitle = (string)(excelWorksheet.Cells[row, 32] as Excel.Range).Value;
                    project.OracleProposalNumber =(string)(excelWorksheet.Cells[row, 33] as Excel.Range).Value;
                    //ABTKID = Project ID
                    project.ProjectsID = (int?)(excelWorksheet.Cells[row, 35] as Excel.Range).Value;
                    project.ProposalName = (string)(excelWorksheet.Cells[row, 34] as Excel.Range).Value;

                }
                catch (Exception e)
                {
                    Program.LogNDisplay("Failed to read Excel. Projects line #" + row + "\n Message: " + e.Message);
                }

                //reconcile against the DB
                if (Program.ProjectsFromDB.ContainsKey(projectNumber))
                {
                    Projects value = Program.ProjectsFromDB[projectNumber];
                    if (project.ProjectNumber.Equals(value.ProjectNumber))
                    {
                        Program.CleanLogNDisplay("Project: " + value.ProjectNumber + " AbtName: " + value.ProjectName + "  #" + count);
                    }
                }
                else
                {

                    Program.LogNDisplay("Key: " + projectNumber + " #" + count + " is not in the dictionary." +
                        "Projects extract line #" + row);
                }
                count++;
                row++;
            }

        }
        public static void LoadSharePointProjectDescriptionsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Sharepoint Project Descriptions \n");

            while ((string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value != null)
            {
                ProjectDescription description = new ProjectDescription();
                string projectNumber = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                try
                {
                    //load row into memory
                    description.ProjectName = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                    description.ProjectNumber = projectNumber;
                    description.Title = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    description.Overview = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                    description.Accomplishments = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                    description.Awards = (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                    description.Innovative = (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                    description.KeyDeliverables = (string)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                    description.Problems = (string)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                    description.ScopeOfWork = (string)(excelWorksheet.Cells[row, 11] as Excel.Range).Value;
                    //if int or string
                    string descType = (string)(excelWorksheet.Cells[row, 12] as Excel.Range).Value;
                    if (descType != null)
                    {
                        if (descType.Equals("Primary"))
                            description.DescriptionType = 1;
                        else
                            description.DescriptionType = 2;
                    }

                    description.DescriptionID = (Int64?)(excelWorksheet.Cells[row, 13] as Excel.Range).Value;
                    //is this abtk id? -> compare against the Overview_ID
                    description.ProjectsID = (int?)(excelWorksheet.Cells[row, 22] as Excel.Range).Value;

                }
                catch (Exception e)
                {
                    Program.LogNDisplay("Failed to read Excel. Project Description line #" + row + "\n Message: " + e.Message);
                }

                //reconcile against the DB
                if (Program.ProjectsFromDB.ContainsKey(projectNumber))
                {
                    try
                    {
                        //find description project

                        //compare
                        Projects value = Program.ProjectsFromDB[projectNumber];
                        if (value.GetDescription(Convert.ToString(description.DescriptionID)).ProjectNumber.Equals(description.ProjectNumber))
                        {
                            Program.CleanLogNDisplay("Project Description: " + description.DescriptionID + "_" + description.Title +
                                " Project Number_Name " + projectNumber + "_" + description.ProjectName + "  #" + count + "in row #" + row);
                        }
                        else
                        {
                            Program.LogNDisplay("Project Description: " + description.DescriptionID + "_" + description.Title + " #" + count + " is not in the dictionary." +
                                "Project Number_Name " + projectNumber + "_" + description.ProjectName + "ProjectDescription extract line #" + row);

                        }
                    }catch (Exception e)
                    {
                        Program.LogNDisplay("Project Description: " + description.DescriptionID + "_" + description.Title + " #" + count + " is not in the dictionary." +
                                                        "Project Number_Name " + projectNumber + "_" + description.ProjectName + "ProjectDescription extract line #" + row);
                    }
                }
                else
                {
                    Program.LogNDisplay("Project number not found: " + projectNumber + " AbtName: " + description.ProjectName + "  #" + count);
                }
                count++;
                row++;
            }
        }
        public static void LoadSharePointPojectsDocumentsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Sharepoint Project Documents \n");

            while ((string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value != null)
            {
                ProjectDocuments document = new ProjectDocuments();
                if((string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value != null)
                {

                    string projectNumber = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    try
                    {
                        //load row into memory
                        document.DocumentName = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                        document.Title = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                        document.ProjectName = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                        document.ProjectNumber = projectNumber;
                        // one of these fields should have an author 
                        if ((string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;

                        }
                        else if ((string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;

                        }
                        else if ((string)(excelWorksheet.Cells[row, 9] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 9] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                        }

                        document.DocumentDate = (DateTime)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                        document.FileSize =Convert.ToString( (double)(excelWorksheet.Cells[row, 21] as Excel.Range).Value);

                    }
                    catch (Exception e)
                    {
                        Program.LogNDisplay("Failed to read Excel. Document in the row #" + row + "\n Message: " + e.Message);
                    }

                    //reconcile against the DB
                    if (Program.ProjectsFromDB.ContainsKey(projectNumber))
                    {
                        try
                        {
                            //find description project
                            if (Program.ProjectsFromDB[projectNumber].DocumentContainsKey(document.DocumentName))
                            {
                                //compare
                                Projects value = Program.ProjectsFromDB[projectNumber];
                                if (value.GetDocuments(document.DocumentName).ProjectNumber != null && 
                                    value.GetDocuments(document.DocumentName).ProjectNumber.Equals(document.ProjectNumber))
                                {
                                    //if (value.GetDocuments(document.DocumentName).FileSize.Equals(document.FileSize)) { }
                                    Program.CleanLogNDisplay("Document Name: " + document.DocumentName + " Project Number_Name: " + document.ProjectNumber + "_" + document.ProjectName + "  #" + count);
                                }
                                else
                                {
                                    Program.LogNDisplay("Document Name: " + document.DocumentName + "for Project Number_Name: " + document.ProjectNumber + "_" + document.ProjectName +
                                        " is not in the dictionary. " + "Review Projects extract in line #" + row);
                                }
                            }
                        }catch(Exception e)
                        {
                            Program.LogNDisplay("Document Name: " + document.DocumentName + "for Project Number_Name: " + document.ProjectNumber + "_" + document.ProjectName +
                                       " is not in the dictionary. " + "Review Projects extract in line #" + row);
                        }
             
                    }
                    else
                    {
                        Program.LogNDisplay(" Project Number_Name not found: " + document.ProjectNumber + "_" + document.ProjectName  + " not found for Document Name: " + document.DocumentName + "  #" + count);
                    }
                }
                else
                {
                    Program.LogNDisplay("Document in row #"+ row +" did not have a project number");
                }



                count++;
                row++;
            }
        }
        public static void LoadSharePointProposalsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Projects \n");

            while ((string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value != null)
            {
                Proposals proposal = new Proposals();
                string proposalNumber = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                int proposalID = (int)(excelWorksheet.Cells[row, 26] as Excel.Range).Value;
                try
                {
                    //load row into memory
                    proposal.ProposalNumber = proposalNumber;
                    proposal.ProposalTitle = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                    proposal.ProposalName = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                    proposal.Comments = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    proposal.Summary = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                    proposal.Client = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                    proposal.IsPrime = (bool)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                    proposal.UltimateClient = (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                    proposal.RFPNumber = (string)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                    proposal.ProjectNumber = (string)(excelWorksheet.Cells[row, 17] as Excel.Range).Value; //17 project oracle number & 19 project number
                    proposal.ProjectName = (string)(excelWorksheet.Cells[row, 18] as Excel.Range).Value;
                    proposal.Division = (string)(excelWorksheet.Cells[row, 20] as Excel.Range).Value;
                    proposal.Practice = (string)(excelWorksheet.Cells[row, 21] as Excel.Range).Value;
                    proposal.ProposalManager = (string)(excelWorksheet.Cells[row, 22] as Excel.Range).Value;
                    proposal.ProposalID = proposalID;
                    //if (((string)(excelWorksheet.Cells[row, 31] as Excel.Range).Value).ToUpper().Equals("YES"))
                    //    proposal.IsActive = true;
                    //else
                    //    proposal.IsActive = false;


                    proposal.FederalAgency = (string)(excelWorksheet.Cells[row, 33] as Excel.Range).Value;
                    proposal.MMG = (string)(excelWorksheet.Cells[row, 34] as Excel.Range).Value;

                }
                catch (Exception e)
                {
                    Program.LogNDisplay("Failed to read Excel. Projects line #" + row + "\n Message: " + e.Message);
                }

                //reconcile against the DB
                if (Program.ProposalsFromDB.ContainsKey(proposal.ProposalNumber))
                {
                    Proposals value = Program.ProposalsFromDB[proposal.ProposalNumber];
                    if (proposal.ProposalNumber.Equals(value.ProposalNumber) &&
                        proposal.ProposalName.Equals(value.ProposalName))
                    {
                        Program.CleanLogNDisplay("Proposal: " + proposal.ProposalNumber + " Proposal Name: " + value.ProposalName + "  #" + count);
                    }
                }
                else
                {

                    Program.LogNDisplay("Key: " + proposalNumber + " #" + count + " is not in the dictionary." +
                        "Projects extract line #" + row);
                }
                count++;
                row++;
            }

        }
        public static void LoadSharePointProposalDocumentsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Sharepoint Project Documents \n");

            while ((string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value != null)
            {
                ProposalDocuments document = new ProposalDocuments();
                if ((string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value != null)
                {

                    string proposalsNumber = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                    try
                    {
                        //load row into memory
                        document.DocumentName = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                        document.ProposalName = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                        document.ProposalNumber = proposalsNumber;
                        // one of these fields should have an author 
                        if ((string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;

                        }
                        else if ((string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;

                        }
                        else if ((string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value != null &&
                            (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value != "")
                        {
                            document.Author = (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                        }

                        document.DocumentDate = (DateTime)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                       // document.FileSize = Convert.ToString((double)(excelWorksheet.Cells[row, 21] as Excel.Range).Value);

                    }
                    catch (Exception e)
                    {
                        Program.LogNDisplay("Failed to read Excel. Document in the row #" + row + "\n Message: " + e.Message);
                    }

                    //reconcile against the DB
                    if (Program.ProposalsFromDB.ContainsKey(proposalsNumber))
                    {
                        try
                        {
                            //find description project
                            if (Program.ProposalsFromDB[proposalsNumber].DocumentContainsKey(document.DocumentName))
                            {
                                //compare
                                Proposals value = Program.ProposalsFromDB[proposalsNumber];
                                if (value.GetDocuments(document.DocumentName).ProposalNumber != null &&
                                    value.GetDocuments(document.DocumentName).ProposalNumber.Equals(document.ProposalNumber))
                                {
                                    //if (value.GetDocuments(document.DocumentName).FileSize.Equals(document.FileSize)) { }
                                    Program.CleanLogNDisplay("Document Name: " + document.DocumentName + " Project Number_Name: " + document.ProposalNumber + "_" + document.ProposalName + "  #" + count);
                                }
                                else
                                {
                                    Program.LogNDisplay("Document Name: " + document.DocumentName + "for Project Number_Name: " + document.ProposalNumber + "_" + document.ProposalName +
                                        " is not in the dictionary. " + "Review Projects extract in line #" + row);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            Program.LogNDisplay("Document Name: " + document.DocumentName + "for Project Number_Name: " + document.ProposalNumber + "_" + document.ProposalNumber +
                                       " is not in the dictionary. " + "Review Projects extract in line #" + row);
                        }

                    }
                    else
                    {
                        Program.LogNDisplay(" Project Number_Name not found: " + document.ProposalNumber + "_" + document.ProposalNumber + " not found for Document Name: " + document.DocumentName + "  #" + count);
                    }
                }
                else
                {
                    Program.LogNDisplay("Document in row #" + row + " did not have a project number");
                }
                count++;
                row++;
            }
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        }

        //excell stuff
        public static void WriteToCell(int row, int column, Worksheet worksheet, string average)
        {
            var cell = (Range)worksheet.Cells[row, column];
            cell.Value2 = average;

        }
        public static void WriteToCell(int row, int column, Worksheet worksheet, double average)
        {
            var cell = (Range)worksheet.Cells[row, column];
            cell.Value2 = average;

        }

    }
}
