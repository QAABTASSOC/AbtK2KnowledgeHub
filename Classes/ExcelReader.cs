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

        public static string[] currentFileNames = new string[14400];
        public static string[] sharepointFileNames = new string[14400];
        public static string[] pathsToSearch = new string[3000];

        public static Dictionary<string, int> indexFinder = new Dictionary<string, int>();

        private static bool start;
        public static bool Start
        {
            get { return start; }
            set { start = value; }
        }

        public static int ReadConfig()
        {
            excelApp = new Excel.Application();
            //TO USE EXCEL: PUT THE PATH IN THE LINE BELOW
            excelWorkbook = excelApp.Workbooks.Open("WHEN USING EXCEL PUT PATH HERE", true, true);
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item("Sheet1");

            LoadCurrentNameAndSharePointName();
            //log its done with names
            Program.LogNDisplay("\n Done Loading Current and Sharedpoint names \n");

            LoadPaths();
            Program.LogNDisplay("\n Done Loading all the posible paths \n");

            excelWorkbook.Close();// change missValue to null
            excelApp.Quit();

            ReleaseObject(excelWorksheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            //returns the drive to be tested
            return 0;
        }
        public static void LoadCurrentNameAndSharePointName()
        {
            int count = 0;
            int row = 2;
            try
            {
                while ((string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value != null)
                {
                    string current = (string)(excelWorksheet.Cells[row, 1] as Excel.Range).Value;
                    string sharepoint_name = "share_point_name";

                    if ((string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value != null)
                    {   //add current name to the array
                        currentFileNames[count] = current;
                        //save new name
                        sharepoint_name = (string)(excelWorksheet.Cells[row, 2] as Excel.Range).Value;
                        //add new name to the array
                        sharepointFileNames[count] = sharepoint_name;
                        //add to index map
                        if (!indexFinder.ContainsKey(current))
                        {
                            indexFinder.Add(current, count);
                            int value = indexFinder[current];
                            Console.WriteLine(value);
                        }
                        else
                        {
                            indexFinder.Add(current, -1);
                            Program.LogNDisplay("the file: " + current + " have been previously processed: " + " \n index: " + count);
                        }

                        //add to log file
                        Program.LogNDisplay("current name: " + current + " SharePoint Name: " + sharepoint_name
                                            + " index: " + count);
                    }
                    else
                    {
                        Program.LogNDisplay("No SharedPoint name found for " + current + " instead: " + sharepoint_name);
                    }
                    count++;
                    row++;
                }
            }
            catch (Exception)
            {

            }

        }
        public static void LoadPaths()
        {
            int count = 0;
            int row = 2;
            try
            {
                while ((string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value != null)
                {
                    string path = (string)(excelWorksheet.Cells[row, 3] as Excel.Range).Value;
                    //add current name to the array
                    pathsToSearch[count] = path;
                    Program.LogNDisplay("Path name: " + path + " index: " + count);

                    count++;
                    row++;
                }
            }
            catch (Exception e)
            {
                Program.LogNDisplay("Error Loading The Paths From Sheet \n\n" + e.Message);
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
