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

        public static List<Projects> ProjectsFromExcel = new List<Projects>();
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
            //TO USE EXCEL: PUT THE PATH IN THE LINE BELOW
            excelWorkbook = excelApp.Workbooks.Open("C:\\Users\\frometaguerraj\\Desktop\\KH_MIGRATION\\"+fileName+".xlsx", true, true);
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Sheet1");

            switch (fileName)
            {
                case "Projects":
                    LoadSharePointPojectsExtract();
                    //Projects are done
                    Program.LogNDisplay("\n Finished Loading Projects from Excel \n");
                    break;
            }

           //LoadPaths();
           // Program.LogNDisplay("\n Done Loading all the posible paths \n");

            excelWorkbook.Close();// change missValue to null
            excelApp.Quit();

            ReleaseObject(excelWorksheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            //returns the drive to be tested
            return 0;
        }
        public static void LoadSharePointPojectsExtract()
        {
            int count = 0;
            int row = 2;
            try
            {
                while ((string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value != null)
                {
                    Projects project = new Projects();
                    string projectNumber = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value; 
                                      
                    if ((string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value != null)
                    {  
                        //load row into memory
                        project.ProjectNumber = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                        project.ProjectTitle = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                        project.ProjectName = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                        project.Division = (string)(excelWorksheet.Cells[row, 4] as Excel.Range).Value;
                        project.Client = (string)(excelWorksheet.Cells[row, 5] as Excel.Range).Value;
                        project.UltimateClient = (string)(excelWorksheet.Cells[row, 6] as Excel.Range).Value;
                        project.IsPrime = (bool)(excelWorksheet.Cells[row, 7] as Excel.Range).Value;
                        project.BeginDate = (DateTime)(excelWorksheet.Cells[row, 8] as Excel.Range).Value;
                        project.EndDate = (DateTime)(excelWorksheet.Cells[row, 9] as Excel.Range).Value;
                        project.OriginalEndDate = (DateTime)(excelWorksheet.Cells[row, 10] as Excel.Range).Value;
                        project.ContractNumber = (string)(excelWorksheet.Cells[row, 11] as Excel.Range).Value;
                        project.AdditionalReference = (string)(excelWorksheet.Cells[row, 12] as Excel.Range).Value;
                        project.PotentialWorth = (Decimal)(excelWorksheet.Cells[row, 13] as Excel.Range).Value;
                        project.AwardAmount= (Decimal)(excelWorksheet.Cells[row, 14] as Excel.Range).Value;
                        project.ProjectType = (string)(excelWorksheet.Cells[row, 15] as Excel.Range).Value;
                        project.FundedAmount = (Decimal)(excelWorksheet.Cells[row, 16] as Excel.Range).Value;
                        project.ProjectDirector = (string)(excelWorksheet.Cells[row, 17] as Excel.Range).Value;
                        project.ProjectDirectorName = (string)(excelWorksheet.Cells[row, 18] as Excel.Range).Value;
                        project.TechnicalOfficer = (string)(excelWorksheet.Cells[row, 19] as Excel.Range).Value;
                        project.TechnicalOfficerName = (string)(excelWorksheet.Cells[row, 20] as Excel.Range).Value;
                        project.Practice = (string)(excelWorksheet.Cells[row, 21] as Excel.Range).Value;
                        project.ProjectStatus = (string)(excelWorksheet.Cells[row, 22] as Excel.Range).Value;
                        project.IsGoodReferenceText= (string)(excelWorksheet.Cells[row, 23] as Excel.Range).Value;
                        project.ParentProject = (int)(excelWorksheet.Cells[row, 24] as Excel.Range).Value;
                        project.IsActive = (bool)(excelWorksheet.Cells[row, 25] as Excel.Range).Value;
                        project.ProjectComments = (string)(excelWorksheet.Cells[row, 26] as Excel.Range).Value;
                        project.AgreementID = (int)(excelWorksheet.Cells[row, 27] as Excel.Range).Value;
                        project.AgreementName = (string)(excelWorksheet.Cells[row, 28] as Excel.Range).Value;
                        project.AgreementTrackNumber = (int)(excelWorksheet.Cells[row, 29] as Excel.Range).Value;
                        project.AgreementType = (string)(excelWorksheet.Cells[row, 30] as Excel.Range).Value;
                        project.FederalAgency = (string)(excelWorksheet.Cells[row, 31] as Excel.Range).Value;
                        project.MMG = (string)(excelWorksheet.Cells[row, 32] as Excel.Range).Value;
                        project.InstClient = (string)(excelWorksheet.Cells[row, 33] as Excel.Range).Value;
                        project.MVTitle = (string)(excelWorksheet.Cells[row, 34] as Excel.Range).Value;
                        project.OracleProposalNumber = (int)(excelWorksheet.Cells[row, 35] as Excel.Range).Value;
                        //is this abtk id? -> compare against the Overview_ID
                        project.ProjectsID = (int)(excelWorksheet.Cells[row, 36] as Excel.Range).Value;

                        //add to index map
                        if (!ExcelProjectsDictionary.ContainsKey(projectNumber))
                        {
                            ExcelProjectsDictionary.Add(projectNumber, project);
                            Projects value = ExcelProjectsDictionary[projectNumber];
                            Console.WriteLine(value.ProjectNumber);
                        }
                        else
                        {
                           // indexFinder.Add(projectNumber, -1);
                            Program.LogNDisplay("the file: " + projectNumber + " have been previously processed: " + " \n index: " + count);
                        }

                        //add to log file
                        Program.LogNDisplay("current number: " + projectNumber + "  #" + count);
                    }
                    else
                    {
                        Program.LogNDisplay("No SharedPoint name found for " + projectNumber + " instead: " );
                    }
                    count++;
                    row++;
                }
            }
            catch (Exception)
            {

            }

        }
        //public static void LoadPaths()
        //{
        //    int count = 0;
        //    int row = 2;
        //    try
        //    {
        //        while ((string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value != null)
        //        {
        //            string path = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
        //            //add current name to the array
        //            pathsToSearch[count] = path;
        //            Program.LogNDisplay("Path name: " + path + " index: " + count);

        //            count++;
        //            row++;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Program.LogNDisplay("Error Loading The Paths From Sheet \n\n" + e.Message);
        //    }


        //}
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
