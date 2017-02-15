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

        private static bool start;
        public static bool Start
        {
            get { return start; }
            set { start = value; }
        }

        public static int ReadConfig(String fileName)
        {
            excelApp = new Excel.Application();
    
            try
            {   //TO USE EXCEL: PUT THE PATH IN THE LINE BELOW
                excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\" + fileName + ".xlsx", true, true);
                excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Sheet1");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
            }         

            switch (fileName)
            {
                case "Projects":
                    LoadSharePointPojectsExtract();
                    //Projects are done
                    Program.LogNDisplay("\n Finished Loading Sharepoint Projects \n");
                    break;
                case "Descriptions":
                    LoadSharePointProjectDescriptionsExtract();
                    //Descriptions are done
                    Program.LogNDisplay("\n Finished Loading Sharepoint Descriptions \n");
                    break;

                case "Documents":
                    LoadSharePointPojectsExtract();
                    //Documents are done
                    Program.LogNDisplay("\n Finished Loading Sharepoint Documents \n");
                    break;

                case "Tags":
                    LoadSharePointPojectsExtract();
                    //Tags are done
                    Program.LogNDisplay("\n Finished Loading Sharepoint Tags \n");
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
                        project.Division = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                        project.Client = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                        project.UltimateClient = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                        project.IsPrime = (bool?)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                        project.BeginDate = (DateTime?)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                        project.EndDate = (DateTime?)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                        project.OriginalEndDate = (DateTime?)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                        project.ContractNumber = (string)(excelWorksheet.Cells[row, 11] as Excel.Range).Value;
                        project.AdditionalReference = (string)(excelWorksheet.Cells[row, 12] as Excel.Range).Value;
                        project.PotentialWorth = (Decimal?)(excelWorksheet.Cells[row, 13] as Excel.Range).Value;
                        project.AwardAmount= (Decimal?)(excelWorksheet.Cells[row, 14] as Excel.Range).Value;
                        project.ProjectType = (string)(excelWorksheet.Cells[row, 15] as Excel.Range).Value;
                        project.FundedAmount = (Decimal?)(excelWorksheet.Cells[row, 16] as Excel.Range).Value;
                        project.ProjectDirector = (string)(excelWorksheet.Cells[row, 17] as Excel.Range).Value;
                        project.ProjectDirectorName = (string)(excelWorksheet.Cells[row, 18] as Excel.Range).Value;
                        project.TechnicalOfficer = (string)(excelWorksheet.Cells[row, 19] as Excel.Range).Value;
                        project.TechnicalOfficerName = (string)(excelWorksheet.Cells[row, 20] as Excel.Range).Value;
                        project.Practice = (string)(excelWorksheet.Cells[row, 21] as Excel.Range).Value;
                        project.ProjectStatus = (string)(excelWorksheet.Cells[row, 22] as Excel.Range).Value;
                        project.IsGoodReferenceText= (string)(excelWorksheet.Cells[row, 23] as Excel.Range).Value;
                        project.ParentProject = (string)(excelWorksheet.Cells[row, 24] as Excel.Range).Value;

                        if (((string)(excelWorksheet.Cells[row, 25] as Excel.Range).Value).ToUpper().Equals("TRUE"))
                        project.IsActive = true;
                        else
                        project.IsActive = false;

                        project.ProjectComments = (string)(excelWorksheet.Cells[row, 26] as Excel.Range).Value;
                        project.AgreementID = (int?)(excelWorksheet.Cells[row, 27] as Excel.Range).Value;
                        project.AgreementName = (string)(excelWorksheet.Cells[row, 28] as Excel.Range).Value;
                        project.AgreementTrackNumber = Convert.ToInt32((string)(excelWorksheet.Cells[row, 29] as Excel.Range).Value);
                        project.AgreementType = (string)(excelWorksheet.Cells[row, 30] as Excel.Range).Value;
                        project.FederalAgency = (string)(excelWorksheet.Cells[row, 31] as Excel.Range).Value;
                        project.MMG = (string)(excelWorksheet.Cells[row, 32] as Excel.Range).Value;
                        project.InstClient = (string)(excelWorksheet.Cells[row, 33] as Excel.Range).Value;
                        project.MVTitle = (string)(excelWorksheet.Cells[row, 34] as Excel.Range).Value;
                        project.OracleProposalNumber = Convert.ToInt32((string)(excelWorksheet.Cells[row, 35] as Excel.Range).Value);
                        //ABTKID = Project ID
                        project.ProjectsID = (int?)(excelWorksheet.Cells[row, 36] as Excel.Range).Value;

                }
                catch (Exception e)
                {
                    Program.LogNDisplay("Failed to read Excel. Projects line #" + row+"\n Message: " + e.Message);
                }

                //reconcile against the DB
                        if (Program.ProjectsFromDB.ContainsKey(projectNumber))
                        {
                            Projects value = Program.ProjectsFromDB[projectNumber];
                            if (project.ProjectNumber.Equals(value.ProjectNumber))
                            {
                                Console.WriteLine("Project: "+value.ProjectNumber+ " AbtName: "+ value.ProjectName + "  #" + count) ;
                            }
                        }
                        else
                        {
                         
                            Program.LogNDisplay("Key: " + projectNumber + " #" + count +" is not in the dictionary."+
                                "Projects extract line #" + row );
                        }
                  count++;
                  row++;
                }
          
        }
        public static void LoadSharePointProjectDescriptionsExtract()
        {
            int count = 1;
            int row = 2;

            Program.LogNDisplay("\n Begin Loading Sharepoint Descriptions \n");

            while ((string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value != null)
            {
                ProjectDescription description = new ProjectDescription();
                string projectNumber = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                try
                {
                    //load row into memory
                    description.ProjectName = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                    description.ProjectNumber = projectNumber;
                    description.Title= (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                    description.Overview = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                    description.Accomplishments = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                    description.Awards = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                    description.Innovative = (string)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                    description.KeyDeliverables = (string)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                    description.Problems = (string)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                    description.ScopeOfWork = (string)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                    //if int or string
                    string descType = (string)(excelWorksheet.Cells[row, 11] as Excel.Range).Value;
                    if (descType != null)
                    {
                        if (descType.Equals( "Primary"))
                            description.DescriptionType = 1;
                        else
                            description.DescriptionType = 2;
                    }
                    
                    description.DescriptionID = (int)(excelWorksheet.Cells[row, 12] as Excel.Range).Value;

                    //need to talk to deepa the sql view cointains projectid not number.
                    // description.ProjectsID = (int?)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;

                    //    if (((string)(excelWorksheet.Cells[row, 25] as Excel.Range).Value).ToUpper().Equals("TRUE"))
                    //{
                    //    description.IsActive = true;
                    //}
                    //else {
                    //    description.IsActive = false;
                    //}
                    //is this abtk id? -> compare against the Overview_ID
                    description.ProjectsID = (int?)(excelWorksheet.Cells[row, 36] as Excel.Range).Value;

                }
                catch (Exception e)
                {
                    Program.LogNDisplay("Failed to read Excel. Description line #" + row + "\n Message: " + e.Message);
                }

                //reconcile against the DB
                if (Program.ProjectsFromDB.ContainsKey(projectNumber)){
                    //find description project
                    if (Program.ProjectsFromDB[projectNumber].DescriptionContainsKey(projectNumber)){
                        //compare
                        Projects value = Program.ProjectsFromDB[projectNumber];
                        if (value.GetDescription(projectNumber).ProjectNumber.Equals(description.ProjectNumber))
                        {
                            Console.WriteLine("Description: " + description.DescriptionID + " AbtName: " + description.ProjectName + "  #" + count);
                        }
                        else
                        {
                            Program.LogNDisplay("Key: " + projectNumber + " #" + count + " is not in the dictionary." +
                                "Projects extract line #" + row);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Project number not found: " + projectNumber + " AbtName: " + description.ProjectName + "  #" + count);
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
