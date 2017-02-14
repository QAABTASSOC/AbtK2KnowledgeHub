using AbtK2KnowledgeHub_OneTime.Classes;
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
                                 + @"\QA_Development\";

        //contains all of the valid formats for the test
        public static HashSet<string> extensions = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase)
        { ".WPD",".wpd",".doc", ".docx", ".xlsx", ".xls",".pdf",".txt", ".pptx",".ppt", ".jpg", ".zip",".png",".rar" };

        public static Dictionary<string, string> ignoredRecords = new Dictionary<string, string>();
        public static Dictionary<string, Projects> ProjetcsFromDB = new Dictionary<string, Projects>();


        static void Main(string[] args)
        {
            Program program = new Program();

            List<AbKLog> logs = new List<AbKLog>();
            try
            {
                if (string.IsNullOrEmpty(program.knowledgeHubWebUrl) || string.IsNullOrEmpty(program.emailId) || string.IsNullOrEmpty(program.password))
                {
                    Console.WriteLine("You have not entered one or more connection parameters. Please check app.config and supply values for weburl, email id and password");
                    return;
                }
               // program.ImportProjects();
              //  program.ApplyTagsOnProject();
              //  program.ApplyTagsOnENRProject();
              //  program.ImportStaff();
                //  program.DuplicateStaffOnENRProjects();
             //   program.ImportDescription();
             //   program.ImportDocuments();
             //   program.ApplyMetaDataToDocuments();
               
                ExcelReader.ReadConfig("Projects");
                program.test();
                Console.WriteLine("Import is complete. Press any key to exit.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                var log = Helper.ConstructLog(Enums.MigrationModule.Other, Enums.LogType.Error, program.correlationId, ex.Message, DateTime.UtcNow);
                logs.Add(log);
            }
            finally
            {
                Helper.WriteLogsToLocalFile(logs, Helper.GeneralLogFilePath);
            }

        }
        //+ " where ProjectNumber in ('17288','19503','20152','19912','18210','16488','07544','17916','18274','19410','16590','20330','20149','20479','13243') " +
        //                    "or ProjectNumber like '17288-%' or ProjectNumber like '19503-%' or ProjectNumber like '20152-%' or ProjectNumber like '19912-%' or " +
        //                    "ProjectNumber like '18210-%' or ProjectNumber like '16488-%' or ProjectNumber like '07544-%' or ProjectNumber like '17916-%'  or " +
        //                    "ProjectNumber like '18274-%'  or ProjectNumber like '19410-%'  or ProjectNumber like '16590-%'  or ProjectNumber like '20330-%' or " +
        //                    "ProjectNumber like '20149-%'  or ProjectNumber like '20479-%'  or ProjectNumber like '13243-%' "
        public void test()
        {
            int countr = 0;

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
                        thisProject.AwardAmount = Helper.SafeGetDecimal(reader, "AwardAmount");
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
                        thisProject.IsPrimeText = (bool)thisProject.IsActive ? "TRUE" : "FALSE";
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
                        if (!ProjetcsFromDB.ContainsKey(projectNumber))
                        {
                            ProjetcsFromDB.Add(projectNumber, thisProject);
                            Projects value = ProjetcsFromDB[projectNumber];
                            Console.WriteLine(value.ProjectNumber);
                        }
                        else
                        {
                            // indexFinder.Add(projectNumber, -1);
                            Program.LogNDisplay("the file: " + projectNumber + " have been previously processed: " + " \n index: " + count);
                        }
                    
                    }
                    sqlConnetion.Close();
                }

                var arrayOfAllKeys = ExcelReader.ExcelProjectsDictionary.Keys.ToArray();
                foreach (var item in arrayOfAllKeys)
                {
                    bool PN = ProjetcsFromDB[item].ProjectNumber.Equals(ExcelReader.ExcelProjectsDictionary[item]);
                    if (PN)
                    {
                        Console.WriteLine(countr+ " Project: " + ProjetcsFromDB[item].ProjectNumber + " have been found" );
                    }
                   
                }
                
            }
        }


        public static void LogNDisplay(string action, long elapsedTipe)
        {
            using (StreamWriter w = System.IO.File.AppendText(logPath + "PerformanceTestLog.txt"))
            {
                Log(action + ": " + TimeSpan.FromMilliseconds(elapsedTipe).ToString(), w);
            }
            using (StreamReader r = System.IO.File.OpenText(logPath + "PerformanceTestLog.txt"))
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
            using (StreamWriter w = System.IO.File.AppendText(logPath + "KH_FILES_AND_PATHS_LOG.txt"))
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
            using (StreamWriter w = System.IO.File.AppendText(logPath + "KH_FILES_AND_PATHS_LOG.txt"))
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
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
        public void ImportProjects()
        {
            AbKLog log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Info, correlationId, Constants.TranformAndPush + " for Project", DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            Console.WriteLine(Constants.TranformAndPush + " for Project");
       
            try
            {
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                    clientContext.Load(projectList);
                    clientContext.ExecuteQuery();

                    using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectViewKey) + " where ProjectNumber in ('17288','19503','20152','19912','18210','16488','07544','17916','18274','19410','16590','20330','20149','20479','13243') " +
                            "or ProjectNumber like '17288-%' or ProjectNumber like '19503-%' or ProjectNumber like '20152-%' or ProjectNumber like '19912-%' or " +
                            "ProjectNumber like '18210-%' or ProjectNumber like '16488-%' or ProjectNumber like '07544-%' or ProjectNumber like '17916-%'  or " +
                            "ProjectNumber like '18274-%'  or ProjectNumber like '19410-%'  or ProjectNumber like '16590-%'  or ProjectNumber like '20330-%' or " +
                            "ProjectNumber like '20149-%'  or ProjectNumber like '20479-%'  or ProjectNumber like '13243-%' ";
                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader();
                            int count = 0; int i = 0;
                            while (reader.Read())
                            {

                                ListItem item = null;
                                int? projectId = 0;
                                Enums.OperationType typeOfOperation = Enums.OperationType.NotKnown;
                                List<AbKLog> recordLogs = new List<AbKLog>();
                                try
                                {
                                    string projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                                    projectId = Helper.SafeGetInt32(reader, "ProjectsID");
                                    if (string.IsNullOrEmpty(projectNumber))
                                        continue;

                                    string query =
                                       @"<View>
                                  <Query>
                                    <Where>
                                      <Eq>
                                        <FieldRef Name='ProjectOracleNumber' />
                                        <Value Type='Text'>" + projectNumber + @"</Value>
                                      </Eq>
                                    </Where>
                                  </Query>
                                  <ViewFields>
                                    <FieldRef Name='ID' />
                                 </ViewFields>
                                </View>";
                                    ListItemCollection itemCollection = GetItems(clientContext, projectList, query);

                                    User ProjectDirector = null; User TechnicalOfficer = null;
                                    string projectDirector = Helper.SafeGetString(reader, "ProjectDirector");
                                    string technicalOfficer = Helper.SafeGetString(reader, "TechnicalOfficer");
                                    //Getting SharePoint user based on email id
                                    ProjectDirector = GetSPUser(clientContext, web, projectDirector);
                                    TechnicalOfficer = GetSPUser(clientContext, web, technicalOfficer);

                                    //Get Parent Project if any
                                    FieldLookupValue parentProjectLookup = null;
                                    string parentProjectNumber = Helper.SafeGetString(reader, "ParentProject");
                                    if (!string.IsNullOrEmpty(parentProjectNumber))
                                    {
                                        string parentProjectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='ProjectOracleNumber' /><Value Type='Text'>" + parentProjectNumber + @"</Value></Eq></Where>
                                                              </Query><ViewFields><FieldRef Name='ID' /></ViewFields></View>";
                                        ListItemCollection parentProjectItemCollection = GetItems(clientContext, projectList, parentProjectCamlQuery);
                                        if (parentProjectItemCollection != null && parentProjectItemCollection.Count > 0)
                                        {
                                            ListItem parentProject = parentProjectItemCollection[0];
                                            parentProjectLookup = new FieldLookupValue();
                                            parentProjectLookup.LookupId = parentProject.Id;
                                        }
                                    }

                                    if (itemCollection != null && itemCollection.Count > 0)
                                    {
                                        item = itemCollection[0];
                                        typeOfOperation = Enums.OperationType.Update;
                                        log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Info, correlationId, Constants.UpdateRecord, DateTime.UtcNow,
                                            projectId.Value, typeOfOperation, item.Id);
                                        recordLogs.Add(log);
                                    }
                                    else
                                    {
                                        typeOfOperation = Enums.OperationType.Add;
                                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                        item = projectList.AddItem(itemCreateInfo);
                                        log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Info, correlationId, Constants.AddingRecord, DateTime.UtcNow,
                                           projectId.Value, typeOfOperation, 0);
                                        recordLogs.Add(log);
                                        item["ProjectOracleNumber"] = projectNumber;
                                        item["AbtkProjectId"] = projectId;
                                    }

                                    string projectTitle = Helper.SafeGetString(reader, "ProjectTitle");
                                    bool? isActive = Helper.SafeGetBool(reader, "IsActive");
                                    bool? isGoodReference = Helper.SafeGetBool(reader, "IsGoodReference");

                                    item["ProjectContractNumber"] = Helper.SafeGetString(reader, "ContractNumber");
                                    item["StartDate"] = Helper.SafeGetDateTime(reader, "BeginDate");
                                    item["_EndDate"] = Helper.SafeGetDateTime(reader, "EndDate");
                                    item["ProjectOriginalEndDate"] = Helper.SafeGetDateTime(reader, "OriginalEndDate");
                                    //May contain upto more than 500 characters so, truncating
                                    item["Title"] = String.IsNullOrEmpty(projectTitle) ? "" : StringExt.Truncate(projectTitle, 255);
                                    item["BS_Project"] = Helper.SafeGetString(reader, "ProjectName");
                                    item["ProjectComments"] = Helper.SafeGetString(reader, "ProjectComments");
                                    item["ProjectPotentialWorth"] = Helper.SafeGetDecimal(reader, "PotentialWorth");
                                    item["ProjectContractValue"] = Helper.SafeGetDecimal(reader, "ContractValue");
                                    item["ProjectCurrentFunding"] = Helper.SafeGetDecimal(reader, "CurrentFunding");
                                    item["ProjectAdditionalReference"] = Helper.SafeGetString(reader, "AdditionalReference");
                                    item["KHIsAbtPrime"] = Convert.ToBoolean(Helper.SafeGetBool(reader, "IsPrime")) ? true : false;
                                    item["KHClient"] = Helper.SafeGetString(reader, "Client");
                                    item["KHUltimateClient"] = Helper.SafeGetString(reader, "UltimateClient");
                                    item["KHAgreementID"] = Helper.SafeGetInt32(reader, "AgreementID");
                                    item["KHAgreementName"] = Helper.SafeGetString(reader, "AgreementName");
                                    item["KHAgreementType"] = Helper.SafeGetString(reader, "AgreementType");
                                    item["KHDivision"] = GetDivision(Helper.SafeGetString(reader, "Division"));
                                    item["KHPractice"] = GetPractice(Helper.SafeGetString(reader, "Practice"));
                                    item["KHInstClient"] = Helper.SafeGetString(reader, "InstClient");
                                    item["KHFederalAgency"] = Helper.SafeGetString(reader, "FederalAgency");
                                    item["KHAgreementTrackNumber"] = Helper.SafeGetDecimal(reader, "AgreementTrackNumber");
                                    item["KHMVTitle"] = Helper.SafeGetString(reader, "MVTitle");
                                    item["KHMMG"] = Helper.SafeGetString(reader, "MMG");
                                    item["ParentProject"] = parentProjectLookup;
                                    item["ProposalOracleNumber"] = Helper.SafeGetString(reader, "Proposalnumber");

                                    if (isGoodReference.HasValue)
                                        item["ProjectIsGoodReference"] = isGoodReference.Value ? "Yes" : "No";
                                    else
                                        item["ProjectIsGoodReference"] = "Unknown";

                                    if (ProjectDirector != null && ProjectDirector.ServerObjectIsNull != null && !ProjectDirector.ServerObjectIsNull.Value)
                                    {
                                        FieldUserValue userValue = new FieldUserValue();
                                        userValue.LookupId = ProjectDirector.Id;
                                        item["BS_ProjectDirector"] = ProjectDirector;//Helper.SafeGetString(reader, "ProjectDirector");
                                    }
                                    else
                                        item["ProjectDirectorOld"] = projectDirector;

                                    if (TechnicalOfficer != null && TechnicalOfficer.ServerObjectIsNull != null && !TechnicalOfficer.ServerObjectIsNull.Value)
                                    {
                                        FieldUserValue userValue = new FieldUserValue();
                                        userValue.LookupId = TechnicalOfficer.Id;
                                        item["ProjectTechnicalOfficer"] = TechnicalOfficer;//Helper.SafeGetString(reader, "ProjectTechnicalOfficer");
                                    }
                                    else
                                        item["ProjectTechnicalOfficerOld"] = technicalOfficer;

                                    item["ProjectStatus"] = Helper.SafeGetString(reader, "ProjectStatus");
                                    item["ProjectType"] = Helper.SafeGetString(reader, "ProjectType");
                                    if (isActive.HasValue)
                                    {
                                        item["Is_x0020_Active"] = isActive.Value ? "Yes" : "No";
                                    }

                                    item.Update();
                                    clientContext.ExecuteQuery();
                                    log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
                                        projectId.Value, typeOfOperation, item.Id);
                                    recordLogs.Add(log);
                                }
                                catch (Exception e)
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Error, correlationId, Constants.ErrorRecordAddedUpdated + " " + e.Message,
                                        DateTime.UtcNow, projectId.Value, typeOfOperation, (item == null || item.ServerObjectIsNull == null || !item.ServerObjectIsNull.Value) ? 0 : item.Id);
                                    recordLogs.Add(log);
                                }
                                finally
                                {
                                    count++; i++;
                                    if (count == 10)
                                    {
                                        Console.WriteLine(String.Format("{0} Projects imported.", i));
                                        count = 0;
                                    }
                                    Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectLogFilePath);
                                }
                            }
                            sqlConnetion.Close();
                            log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Info, correlationId, "Project import is complete.", DateTime.UtcNow);
                            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
                            Console.WriteLine("Projects Imported.");
                        }
                    }
                }
            }
            catch (WebException we)
            {
                Console.WriteLine("Some error has occured while connecting to SharePoint Site: " + we.Message);
                log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Error, correlationId, we.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Project, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
        }

        public void ImportStaff()
        {
            AbKLog log;
            log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Info, correlationId, Constants.TranformAndPush + " for Staff.", DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
            Console.WriteLine(Constants.TranformAndPush + " for Staff.");
            try
            {
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                    clientContext.Load(projectList);
                    List projectStaffList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectStaffListKey));
                    clientContext.Load(projectStaffList);
                    clientContext.ExecuteQuery();

                    using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectStaffViewKey);

                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader(); int i = 0; int count = 0;

                            int lastProjectId = 0; ListItem project = null;
                            while (reader.Read())
                            {
                                Enums.OperationType typeOfOperation = Enums.OperationType.NotKnown;
                                List<AbKLog> recordLogs = new List<AbKLog>();
                                ListItem employeeItem = null;
                                int? staffRecordId = 0;
                                try
                                {
                                    int? projectsID = Helper.SafeGetInt32(reader, "ProjectsID");
                                    string emailID = Helper.SafeGetString(reader, "Email");
                                    string projectStaffFullName = Helper.SafeGetString(reader, "ProjectStaffFullName");
                                    string employeeRole = Helper.SafeGetString(reader, "StaffRole");

                                    staffRecordId = Helper.SafeGetInt32(reader, "ProjectStaffID");
                                    if (!projectsID.HasValue)
                                        continue;
                                    if (string.IsNullOrEmpty(emailID) && string.IsNullOrEmpty(projectStaffFullName))
                                        continue;
                                    if (string.IsNullOrEmpty(employeeRole))
                                        continue;
                                    if (lastProjectId != projectsID.Value)
                                    {
                                        string projectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='AbtkProjectId' /><Value Type='Text'>" + projectsID + @"</Value></Eq></Where>
                                                              </Query><ViewFields><FieldRef Name='ID' /></ViewFields></View>";
                                        ListItemCollection projectListItemCollection = GetItems(clientContext, projectList, projectCamlQuery);
                                        if (projectListItemCollection != null && projectListItemCollection.Count > 0)
                                            project = projectListItemCollection[0];
                                        else
                                        {
                                            log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Error, correlationId, "Project not found",
                                                DateTime.UtcNow, staffRecordId.Value, typeOfOperation, 0);
                                            recordLogs.Add(log);
                                            continue;
                                        }
                                    }
                                    lastProjectId = projectsID.Value;


                                    string itemCamlQuery = "";

                                    User employee = GetSPUser(clientContext, web, emailID);

                                    //if (employee != null && employee.ServerObjectIsNull.HasValue && !employee.ServerObjectIsNull.Value)
                                    //{
                                    //    itemCamlQuery = @"<View><Query><Where><And><And><Eq><FieldRef Name='KHProject' LookupId='True'/><Value Type='Lookup'>" + project.Id +
                                    //                       @"</Value></Eq><Eq><FieldRef Name='KH_Employee' LookupId='TRUE' /><Value Type='Lookup'>" + employee.Id +
                                    //                       @"</Value></Eq></And><Eq><FieldRef Name='BS_Role'/><Value Type='Text'>" + employeeRole +
                                    //                       @"</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields></View>";
                                    //}
                                    //else if (!string.IsNullOrEmpty(projectStaffFullName))
                                    //{
                                    //    itemCamlQuery = @"<View><Query><Where><And><And><Eq><FieldRef Name='KHProject' LookupId='True'/><Value Type='Lookup'>" + project.Id +
                                    //                       @"</Value></Eq><Eq><FieldRef Name='KHEmployeeName' /><Value Type='Text'>" + projectStaffFullName +
                                    //                       @"</Value></Eq></And><Eq><FieldRef Name='BS_Role'/><Value Type='Text'>" + employeeRole +
                                    //                       @"</Value></Eq></And></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields></View>";
                                    //}
                                    //else
                                    //{
                                    //    continue;
                                    //}
                                    itemCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='AbtkProjectStaffID' /><Value Type='Text'>" + staffRecordId.Value +
                                                           @"</Value></Eq></Where></Query><ViewFields><FieldRef Name='ID' /></ViewFields></View>";
                                    ListItemCollection staffItemCollection = GetItems(clientContext, projectStaffList, itemCamlQuery);

                                    if (staffItemCollection != null && staffItemCollection.Count > 0)
                                    {
                                        employeeItem = staffItemCollection[0];
                                        typeOfOperation = Enums.OperationType.Update;
                                        log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Info, correlationId, Constants.UpdateRecord, DateTime.UtcNow,
                                            staffRecordId.Value, typeOfOperation, employeeItem.Id);
                                        recordLogs.Add(log);
                                    }
                                    else
                                    {
                                        typeOfOperation = Enums.OperationType.Add;
                                        log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Info, correlationId, Constants.AddingRecord, DateTime.UtcNow,
                                          staffRecordId.Value, typeOfOperation, 0);
                                        recordLogs.Add(log);
                                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                        employeeItem = projectStaffList.AddItem(itemCreateInfo);
                                        employeeItem["AbtkProjectStaffID"] = staffRecordId;
                                    }
                                    FieldLookupValue projectLookup = new FieldLookupValue();
                                    projectLookup.LookupId = project.Id;

                                    employeeItem["KHProject"] = projectLookup;
                                    employeeItem["ProjectNumber"] = projectLookup;

                                    if (employee != null && employee.ServerObjectIsNull != null && !employee.ServerObjectIsNull.Value)
                                    {
                                        FieldUserValue userValue = new FieldUserValue();
                                        userValue.LookupId = employee.Id;
                                        employeeItem["KH_Employee"] = userValue;
                                    }
                                    else
                                    {
                                        employeeItem["KHEmployeeName"] = projectStaffFullName;
                                    }
                                    employeeItem["BS_Role"] = employeeRole;
                                    employeeItem["StartDate"] = Helper.SafeGetDateTime(reader, "StartDate");
                                    employeeItem["_EndDate"] = Helper.SafeGetDateTime(reader, "EndDate");

                                    employeeItem.Update();
                                    clientContext.ExecuteQuery();
                                    log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
                                        staffRecordId.Value, typeOfOperation, employeeItem.Id);
                                    recordLogs.Add(log);

                                }
                                catch (Exception e)
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Error, correlationId, Constants.ErrorRecordAddedUpdated + " " + e.Message,
                                        DateTime.UtcNow, staffRecordId.Value, typeOfOperation, (employeeItem != null && employeeItem.ServerObjectIsNull != null && employeeItem.ServerObjectIsNull.Value) ? employeeItem.Id : 0);
                                    recordLogs.Add(log);
                                }
                                finally
                                {
                                    count++; i++;
                                    if (count == 10)
                                    {
                                        Console.WriteLine(String.Format("{0} Staff imported.", i));
                                        count = 0;
                                    }
                                    Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectStaffLogFilePath);
                                }
                            }
                            sqlConnetion.Close();
                            log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Info, correlationId, "Staff import is complete.", DateTime.UtcNow);
                            Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
                            Console.WriteLine("Staff Imported.");
                        }
                    }
                }

            }
            catch (WebException we)
            {
                Console.WriteLine("Some error has occured while connecting to SharePoint Site: " + we.Message);
                log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Error, correlationId, we.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Staff, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
            }
        }

        public void DuplicateStaffOnENRProjects()
        {
            Enums.OperationType typeOfOperation = Enums.OperationType.Add;
            var log = Helper.ConstructLog(Enums.MigrationModule.ENRStaff, Enums.LogType.Info, correlationId, "Adding staff for ENR tasks.", DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);

            Console.WriteLine("Adding staff for ENR tasks...");
            try
            {
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                    clientContext.Load(projectList);
                    List staffList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectStaffListKey));
                    clientContext.Load(staffList);
                    string query = @"<View><Query><Where><IsNotNull><FieldRef Name='ParentProject' /></IsNotNull></Where><OrderBy>
                                      <FieldRef Name='ParentProject' Ascending='TRUE' LookupId='TRUE' /></OrderBy></Query><ViewFields><FieldRef Name='ParentProject' /><FieldRef Name='Id' />
                                     </ViewFields></View>";
                    ListItemCollection enrProjectCollection = GetItems(clientContext, projectList, query);
                    if (enrProjectCollection == null || enrProjectCollection.ServerObjectIsNull == null || enrProjectCollection.Count == 0)
                        return;
                    int i = 0; int count = 0; int lastProjectID = 0;
                    ListItemCollection parentStaffCollection = null;
                    foreach (ListItem enrItem in enrProjectCollection)
                    {

                        List<AbKLog> recordLogs = new List<AbKLog>();
                        try
                        {
                            FieldLookupValue parentProjectLookup = enrItem["ParentProject"] as FieldLookupValue;
                            if (parentProjectLookup == null)
                                continue;
                            if (parentProjectLookup.LookupId != lastProjectID)
                            {
                                string parentStaffQuery = @"<View><Query><Where><Eq><FieldRef Name='KHProject' LookupId='True' /><Value Type='Lookup'>" + parentProjectLookup.LookupId +
                                                    @"</Value></Eq></Where></Query><ViewFields><FieldRef Name='KHProject' /><FieldRef Name='ProjectNumber' />" +
                                                    @"<FieldRef Name='KH_Employee' /><FieldRef Name='BS_Role' /><FieldRef Name='StartDate' /><FieldRef Name='_EndDate' /><FieldRef Name='Id' />" +
                                                    @"</ViewFields></View>";
                                parentStaffCollection = GetItems(clientContext, staffList, parentStaffQuery);
                            }
                            lastProjectID = parentProjectLookup.LookupId;
                            string childStaffQuery = @"<View><Query><Where><Eq><FieldRef Name='KHProject' LookupId='True' /><Value Type='Lookup'>" + enrItem.Id +
                                            @"</Value></Eq></Where></Query></View>";
                            ListItemCollection childStaffCollection = GetItems(clientContext, staffList, childStaffQuery);

                            if (childStaffCollection != null)
                            {
                                foreach (ListItem childItem in childStaffCollection)
                                {
                                    ListItem item = childStaffCollection.GetById(childItem.Id);
                                    item.DeleteObject();
                                    clientContext.ExecuteQuery();
                                }
                            }

                            if (parentStaffCollection != null || parentStaffCollection.Count >= 0)
                            {
                                foreach (ListItem parentItem in parentStaffCollection)
                                {
                                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                    ListItem employeeItem = staffList.AddItem(itemCreateInfo);
                                    FieldLookupValue projectLookup = new FieldLookupValue();
                                    projectLookup.LookupId = enrItem.Id;
                                    employeeItem["KHProject"] = projectLookup;
                                    employeeItem["ProjectNumber"] = projectLookup;
                                    employeeItem["KH_Employee"] = parentItem["KH_Employee"];
                                    employeeItem["BS_Role"] = parentItem["BS_Role"];
                                    employeeItem["StartDate"] = parentItem["StartDate"];
                                    employeeItem["_EndDate"] = parentItem["_EndDate"];
                                    employeeItem.Update();
                                    clientContext.ExecuteQuery();
                                    log = Helper.ConstructLog(Enums.MigrationModule.ENRStaff, Enums.LogType.Info, correlationId, "ENR Staff added.",
                                DateTime.UtcNow, parentItem.Id, typeOfOperation, employeeItem.Id);
                                    recordLogs.Add(log);
                                }
                            }

                        }
                        catch (Exception e)
                        {
                            log = Helper.ConstructLog(Enums.MigrationModule.ENRStaff, Enums.LogType.Error, correlationId, Constants.ErrorRecordAddedUpdated + " " + e.Message,
                                DateTime.UtcNow, enrItem.Id, typeOfOperation, 0);
                            recordLogs.Add(log);
                        }
                        finally
                        {
                            i++; count++;
                            if (count == 10)
                            {
                                Console.WriteLine(String.Format("Staff members for {0} ENR Record has been added.", i));
                                count = 0;
                            }
                            Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectStaffLogFilePath);
                        }
                    }
                    log = Helper.ConstructLog(Enums.MigrationModule.ENRStaff, Enums.LogType.Info, correlationId, "Staff for ENR Staff has been added.", DateTime.UtcNow);
                    Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
                    Console.WriteLine("Staff for ENR Staff has been added.");
                }
            }
            catch (Exception ex)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.ENRStaff, Enums.LogType.Error, correlationId, ex.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectStaffLogFilePath);
            }
        }

        public void ImportDescription()
        {
            AbKLog log;
            log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Info, correlationId, Constants.TranformAndPush + " for description and adding tags to description.", DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            Console.WriteLine(Constants.TranformAndPush + " for description and adding tags to description.");
            try
            {
                
                    using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectDescriptionViewKey) +
                            "  where ProjectNumber in ('17288','19503','20152','19912','18210','16488','07544','17916','18274','19410','16590','20330','20149','20479','13243') " +
                            "or ProjectNumber like '17288-%' or ProjectNumber like '19503-%' or ProjectNumber like '20152-%' or ProjectNumber like '19912-%' or " +
                            "ProjectNumber like '18210-%' or ProjectNumber like '16488-%' or ProjectNumber like '07544-%' or ProjectNumber like '17916-%'  or " +
                            "ProjectNumber like '18274-%'  or ProjectNumber like '19410-%'  or ProjectNumber like '16590-%'  or ProjectNumber like '20330-%' or " +
                            "ProjectNumber like '20149-%'  or ProjectNumber like '20479-%'  or ProjectNumber like '13243-%' ";

                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader();
                            int count = 0; int i = 0;
                            ListItem project = null;
                            string lastProjectId = string.Empty;
                            while (reader.Read())
                            {
                                Enums.OperationType typeOfOperation = Enums.OperationType.NotKnown;
                                List<AbKLog> recordLogs = new List<AbKLog>();
                                ListItem descriptionItem = null;
                                Int64? descriptionRecordId = 0;
                                try
                                {
                                    string projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                                    string descriptionTitle = Helper.SafeGetString(reader, "Title");
                                    //unique id field in the view in sql
                                    descriptionRecordId = Helper.SafeGetInt64(reader, "OverviewID");
                                    int? descriptionType = Helper.SafeGetInt32(reader, "DescriptionType");
                                    if (string.IsNullOrEmpty(projectNumber) || !descriptionRecordId.HasValue)
                                        continue;
                                   
                                    
                                }
                                catch (Exception e)
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Error, correlationId, Constants.ErrorRecordAddedUpdated + " " + e.Message,
                                        DateTime.UtcNow, descriptionRecordId.Value, typeOfOperation, (descriptionItem != null && descriptionItem.ServerObjectIsNull != null && descriptionItem.ServerObjectIsNull.Value) ? descriptionItem.Id : 0);
                                    recordLogs.Add(log);
                                }
                                finally
                                {
                                    i++; count++;
                                    if (count == 10)
                                    {
                                        Console.WriteLine(String.Format("{0} Descriptions imported.", i));
                                        count = 0;
                                    }
                                    Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectDescriptionLogFilePath);

                                }
                            }
                            sqlConnetion.Close();
                          //  log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Info, correlationId, "Description import is complete.", DateTime.UtcNow);
                         //   Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
                            Console.WriteLine("Description Loaded");
                        }
                    }
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
        }

        public void ApplyTagsOnProject()
        {

            AbKLog log;
            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, Constants.ApplyingTags + " for Project.", DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            Console.WriteLine(Constants.ApplyingTags + " for Project.");

            try
            {
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);

                    using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectTagsViewKey);

                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader();
                            if (!reader.HasRows)
                                return;

                            ListItem project = null;
                            int? tagID = 0;
                            int? projectsID = 0;
                            int? lastProjectsID = 0;

                            List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                            clientContext.Load(projectList);
                            List projectTagsList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.TagListKey));
                            clientContext.Load(projectTagsList);
                            clientContext.ExecuteQuery();
                            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                            clientContext.Load(taxonomySession);
                            clientContext.Load(termStore);
                            clientContext.ExecuteQuery();

                            string query = @"<View><ViewFields><FieldRef Name='Title' /><FieldRef Name='SharePoint_x0020_Term_x0020_Grou' />
                                                            <FieldRef Name='SharePointTermSet' /><FieldRef Name='SharePoint_x0020_Term' /></ViewFields></View>";
                            ListItemCollection tagsCollection = GetItems(clientContext, projectTagsList, query);
                            int i = 0; int count = 0;
                            while (reader.Read())
                            {
                                Enums.OperationType typeOfOperation = Enums.OperationType.ApplyingTags;
                                List<AbKLog> recordLogs = new List<AbKLog>();
                                TaxonomyFieldValueCollection currentCollection = null;
                                string termValueString = string.Empty;
                                try
                                {
                                    projectsID = Helper.SafeGetInt32(reader, "ProjectsID");
                                    tagID = Helper.SafeGetInt32(reader, "ProjectTagID");
                                    if (!projectsID.HasValue || !tagID.HasValue)
                                        continue;

                                    ListItem tagItem = tagsCollection.Where(tag => tag["Title"].ToString() == tagID.ToString()).First();
                                    if (tagItem["SharePoint_x0020_Term"] == null)
                                    {
                                        log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, "Term not found.",
                                                DateTime.UtcNow, tagID.Value, typeOfOperation, projectsID.Value);
                                        recordLogs.Add(log);
                                        continue;
                                    }

                                    Term term = termStore.GetTerm(new Guid(Convert.ToString(tagItem["SharePoint_x0020_Term"])));
                                    clientContext.Load(term, t => t.Id, t => t.Name, t => t.TermSet.Id);
                                    clientContext.ExecuteQuery();
                                    clientContext.Load(term.TermSet, t => t.Name, t => t.Id);
                                    clientContext.ExecuteQuery();
                                    if (!lastProjectsID.HasValue || lastProjectsID.Value != projectsID.Value)
                                    {
                                        string projectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='AbtkProjectId' /><Value Type='Text'>" + projectsID + @"</Value></Eq>
                                                                    </Where></Query><ViewFields><FieldRef Name='ID' /><FieldRef Name='globalAbtCapabilities' />
                                                                    <FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />
                                                                    <FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
                                        ListItemCollection projectListItemCollection = GetItems(clientContext, projectList, projectCamlQuery);

                                        if (projectListItemCollection != null && projectListItemCollection.Count > 0)
                                            project = projectListItemCollection[0];
                                        else
                                        {
                                            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, "Project not found.",
                                                DateTime.UtcNow, tagID.Value, typeOfOperation, projectsID.Value);
                                            recordLogs.Add(log);
                                            continue;
                                        }
                                    }
                                    lastProjectsID = projectsID;
                                    Field field = null; TaxonomyField txField = null;
                                    switch (term.TermSet.Id.ToString())
                                    {
                                        case "2d495acb-97e2-490f-85eb-2f9c5dca2133": //capabilities
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalAbtCapabilities");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalAbtCapabilities"] as TaxonomyFieldValueCollection;
                                            break;
                                        case "c72a75cc-1211-4971-9e55-eafa85bd7e55": // client types
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalClientTypes");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalClientTypes"] as TaxonomyFieldValueCollection;
                                            break;
                                        case "d274570b-8632-492b-be2a-35b05f23b980": //demographics
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalProjectDemographics");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalProjectDemographics"] as TaxonomyFieldValueCollection;
                                            break;
                                        case "33c5bb3c-5cef-4220-8a60-7451d2445763": // geographic locations
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalGeographicLocations");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalGeographicLocations"] as TaxonomyFieldValueCollection;
                                            break;
                                        case "2f488bf4-6e7e-4703-b449-0b2c443e6e4f": //organization
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalAbtOrganization");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalAbtOrganization"] as TaxonomyFieldValueCollection;
                                            break;
                                        case "59c412eb-f2b1-41ac-ba78-b2342d4e8a66": //subject mater areas
                                            field = projectList.Fields.GetByInternalNameOrTitle("globalSubjectMatterAreas");
                                            txField = clientContext.CastTo<TaxonomyField>(field);
                                            currentCollection = project["globalSubjectMatterAreas"] as TaxonomyFieldValueCollection;
                                            break;
                                    }

                                    if (txField == null || currentCollection == null || currentCollection.ServerObjectIsNull == null) continue;

                                    TaxonomyFieldValue value = currentCollection.Where(tvc => tvc.TermGuid == term.Id.ToString()).FirstOrDefault();
                                    if (value == null)
                                    {
                                        termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                        if (termValueString.Length > 0)
                                            termValueString += ";#-1;#" + term.Name + "|" + term.Id;
                                        else
                                            termValueString += "-1;#" + term.Name + "|" + term.Id;
                                        currentCollection = new TaxonomyFieldValueCollection(clientContext, termValueString, txField);
                                        txField.SetFieldValueByValueCollection(project, currentCollection);
                                        project.Update();
                                        clientContext.ExecuteQuery();
                                        log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
                                       tagID.Value, typeOfOperation, projectsID.Value);
                                        recordLogs.Add(log);
                                    }

                                }
                                catch (Exception e)
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, Constants.ErrorInApplyingTags + " " + e.Message,
                                        DateTime.UtcNow, tagID.Value, typeOfOperation, projectsID.Value);
                                    recordLogs.Add(log);

                                }
                                finally
                                {
                                    i++; count++;
                                    if (count == 10)
                                    {
                                        Console.WriteLine(String.Format("{0} tags applied.", i));
                                        count = 0;
                                    }
                                    Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectLogFilePath);
                                }
                            }
                            sqlConnetion.Close();
                            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, "Tags has been applied on Projects.", DateTime.UtcNow);
                            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
                            Console.WriteLine("Tags applied on Projects.");
                        }
                    }
                }

            }
            catch (WebException we)
            {
                Console.WriteLine("Some error has occured while connecting to SharePoint Site: " + we.Message);
                log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Error, correlationId, we.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
        }

        //public void ApplyTagsOnDescription()
        //{
        //    List<AbKLog> generalLogs = new List<AbKLog>();
        //    Enums.OperationType typeOfOperation = Enums.OperationType.ApplyingTags;
        //    var log = Helper.ConstructLog(Enums.MigrationModule.Description, Enums.LogType.Info, correlationId, Constants.ApplyingTags, DateTime.UtcNow);
        //    generalLogs.Add(log);
        //    Console.WriteLine(Constants.ApplyingTags + " on Description.");
        //    try
        //    {
        //        using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
        //        {
        //            clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
        //            Web web = clientContext.Web;
        //            clientContext.Load(web);
        //            List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
        //            clientContext.Load(projectList);
        //            List projectDescriptionList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectDescriptionListKey));
        //            clientContext.Load(projectDescriptionList);
        //            clientContext.ExecuteQuery();
        //            string query = @"<View><Query><OrderBy><FieldRef Name='KHProject' Ascending='TRUE' LookupId='TRUE' /></OrderBy></Query>" +
        //                            @"<ViewFields><FieldRef Name='KHProject' /><FieldRef Name='Id' /><FieldRef Name='globalAbtCapabilities' />
        //                                                            <FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />
        //                                                            <FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
        //            ListItemCollection descriptionCollection = GetItems(clientContext, projectDescriptionList, query);
        //            if (descriptionCollection == null || descriptionCollection.ServerObjectIsNull == null || descriptionCollection.Count == 0)
        //                return;
        //            int i = 0; int count = 0; int lastProjectId = 0;
        //            ListItem project = null;
        //            foreach (ListItem descriptionItem in descriptionCollection)
        //            {
        //                List<AbKLog> recordLogs = new List<AbKLog>();
        //                string termValueString = string.Empty;
        //                try
        //                {
        //                    FieldLookupValue projectLookup = descriptionItem["KHProject"] as FieldLookupValue;
        //                    if (projectLookup == null)
        //                        continue;
        //                    if (lastProjectId != projectLookup.LookupId)
        //                    {
        //                        string projectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + projectLookup.LookupId + @"</Value></Eq>
        //                                                            </Where></Query><ViewFields><FieldRef Name='globalAbtCapabilities' />
        //                                                            <FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />
        //                                                            <FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
        //                        ListItemCollection projectListItemCollection = GetItems(clientContext, projectList, projectCamlQuery);
        //                        if (projectListItemCollection != null && projectListItemCollection.Count > 0)
        //                            project = projectListItemCollection[0];
        //                        else
        //                        {
        //                            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, "Project not found.",
        //                                DateTime.UtcNow, descriptionItem.Id, typeOfOperation, projectLookup.LookupId);
        //                            recordLogs.Add(log);
        //                            continue;
        //                        }
        //                    }
        //                    lastProjectId = projectLookup.LookupId;

        //                    TaxonomyFieldValueCollection currentCollection = project["globalAbtCapabilities"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalAbtCapabilities");
        //                    currentCollection = project["globalAbtOrganization"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalAbtOrganization");
        //                    currentCollection = project["globalClientTypes"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalClientTypes");
        //                    currentCollection = project["globalGeographicLocations"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalGeographicLocations");
        //                    currentCollection = project["globalProjectDemographics"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalProjectDemographics");
        //                    currentCollection = project["globalSubjectMatterAreas"] as TaxonomyFieldValueCollection;
        //                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
        //                    if (!string.IsNullOrEmpty(termValueString))
        //                        SetTaxonomyCollection(clientContext, projectDescriptionList, termValueString, descriptionItem, "globalSubjectMatterAreas");
        //                    descriptionItem.Update();
        //                    clientContext.ExecuteQuery();
        //                    log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
        //                    project.Id, typeOfOperation, descriptionItem.Id);
        //                    recordLogs.Add(log);

        //                }
        //                catch (Exception e)
        //                {
        //                    log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, Constants.ErrorInApplyingTags + " " + e.Message,
        //                        DateTime.UtcNow, project.Id, typeOfOperation, descriptionItem.Id);
        //                    recordLogs.Add(log);
        //                }
        //                finally
        //                {
        //                    i++; count++;
        //                    if (count == 10)
        //                    {
        //                        Console.WriteLine(String.Format("{0} tags applied.", i));
        //                        count = 0;
        //                    }
        //                    Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectDescriptionLogFilePath);
        //                }
        //            }
        //            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, "Tags has been applied on Descriptions.", DateTime.UtcNow);
        //            generalLogs.Add(log);
        //            Console.WriteLine("Tags applied on Descriptions.");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log = Helper.ConstructLog(Enums.MigrationModule.Other, Enums.LogType.Error, correlationId, ex.Message, DateTime.UtcNow);
        //        generalLogs.Add(log);
        //    }
        //    finally
        //    {
        //        Helper.WriteLogsToLocalFile(generalLogs, Helper.ProjectDescriptionLogFilePath);
        //    }
        //}

        public void ApplyTagsOnENRProject()
        {
            Enums.OperationType typeOfOperation = Enums.OperationType.ApplyingTags;
            var log = Helper.ConstructLog(Enums.MigrationModule.ENRTasks, Enums.LogType.Info, correlationId, Constants.ApplyingTags, DateTime.UtcNow);
            Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            Console.WriteLine(Constants.ApplyingTags + " on ENR Task Project.");
            try
            {
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                    clientContext.Load(projectList);

                    string query = @"<View><Query><Where><IsNotNull><FieldRef Name='ParentProject' /></IsNotNull></Where>" +
                                   @"<OrderBy><FieldRef Name='ParentProject' Ascending='TRUE' LookupValue='TRUE' /></OrderBy>" +
                                   @"</Query><ViewFields><FieldRef Name='ParentProject' /><FieldRef Name='Id' /><FieldRef Name='globalAbtCapabilities' />" +
                                    @"<FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />" +
                                  @"<FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
                    ListItemCollection enrProjectCollection = GetItems(clientContext, projectList, query);
                    if (enrProjectCollection == null || enrProjectCollection.ServerObjectIsNull == null || enrProjectCollection.Count == 0)
                        return;
                    int i = 0; int count = 0; ListItem project = null; int lastProjectId = 0;
                    foreach (ListItem enrItem in enrProjectCollection)
                    {
                        List<AbKLog> recordLogs = new List<AbKLog>();
                        string termValueString = string.Empty;
                        try
                        {
                            FieldLookupValue parentProjectLookup = enrItem["ParentProject"] as FieldLookupValue;
                            if (parentProjectLookup == null)
                                continue;
                            if (lastProjectId != parentProjectLookup.LookupId)
                            {
                                string projectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + parentProjectLookup.LookupId + @"</Value></Eq>
                                                                    </Where></Query><ViewFields><FieldRef Name='globalAbtCapabilities' />
                                                                    <FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />
                                                                    <FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
                                ListItemCollection projectListItemCollection = GetItems(clientContext, projectList, projectCamlQuery);
                                if (projectListItemCollection != null && projectListItemCollection.Count > 0)
                                    project = projectListItemCollection[0];
                                else
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.ENRTasks, Enums.LogType.Error, correlationId, "Project not found, while applying tags.",
                                        DateTime.UtcNow, enrItem.Id, typeOfOperation, parentProjectLookup.LookupId);
                                    recordLogs.Add(log);
                                    continue;
                                }
                            }
                            lastProjectId = parentProjectLookup.LookupId;

                            TaxonomyFieldValueCollection currentCollection = project["globalAbtCapabilities"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalAbtCapabilities");
                            currentCollection = project["globalAbtOrganization"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalAbtOrganization");
                            currentCollection = project["globalClientTypes"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalClientTypes");
                            currentCollection = project["globalGeographicLocations"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalGeographicLocations");
                            currentCollection = project["globalProjectDemographics"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalProjectDemographics");
                            currentCollection = project["globalSubjectMatterAreas"] as TaxonomyFieldValueCollection;
                            termValueString = GetTaxonomyStringFromCollection(currentCollection);
                            if (!string.IsNullOrEmpty(termValueString))
                                SetTaxonomyCollection(clientContext, projectList, termValueString, enrItem, "globalSubjectMatterAreas");
                            enrItem.Update();
                            clientContext.ExecuteQuery();
                            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
                            project.Id, typeOfOperation, enrItem.Id);
                            recordLogs.Add(log);
                        }
                        catch (Exception e)
                        {
                            log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, Constants.ErrorInApplyingTags + " " + e.Message,
                                DateTime.UtcNow, project.Id, typeOfOperation, enrItem.Id);
                            recordLogs.Add(log);
                        }
                        finally
                        {
                            i++; count++;
                            if (count == 10)
                            {
                                Console.WriteLine(String.Format("{0} tags been applied on ENR Task Project.", i));
                                count = 0;
                            }
                            Helper.WriteLogsToLocalFile(recordLogs, Helper.ProjectLogFilePath);
                        }
                    }
                    log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Info, correlationId, "Tags has been applied on ENR Task Project.", DateTime.UtcNow);
                    Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
                    Console.WriteLine("Tags applied on ENR Task Project");
                }
            }
            catch (Exception ex)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Tags, Enums.LogType.Error, correlationId, ex.Message, DateTime.UtcNow);
                Helper.WriteLogToLocalFile(log, Helper.ProjectLogFilePath);
            }
        }

        public void ImportDocuments()
        {

            AbKLog log;
            log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Info, correlationId, "Uploading Project documents.", DateTime.UtcNow);
            Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);

            Console.WriteLine("Uploading Project documents.");
            try
            {
                string directoryPath = Helper.GetAppSettingValue(Constants.DocumentsDirectoryPath);
                string documentsListName = Helper.GetAppSettingValue(Constants.DocumentListKey);
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    int i = 0; int count = 0;
                    using (OleDbConnection connection = new OleDbConnection(Helper.GetConnectionString(Constants.ExcelConnectionStringKey)))
                    {
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        adapter.SelectCommand = new OleDbCommand("Select * from [Sheet1$] where Status is Null or Status='0'", connection);
                        DataSet documentMetadata = new DataSet();
                        adapter.Fill(documentMetadata);
                        connection.Open();
                        OleDbCommand updateCommand = new OleDbCommand();
                        updateCommand.Connection = connection;
                        foreach (DataRow pRow in documentMetadata.Tables[0].Rows)
                        {
                            string fileName = Convert.ToString(pRow["UploadedFileName"]);
                            if (string.IsNullOrEmpty(fileName))
                                continue;

                            string filePath = Directory.GetFiles(directoryPath, fileName).FirstOrDefault();
                            string updateCommandString = string.Empty;
                            List<AbKLog> recordLogs = new List<AbKLog>();
                            try
                            {
                                if (!string.IsNullOrEmpty(filePath))
                                {
                                    Microsoft.SharePoint.Client.File item = UploadFileSlicePerSlice(clientContext, documentsListName, filePath, 10);

                                    updateCommandString = String.Format("UPDATE [Sheet1$] Set Status={0} where UploadedFileName='{1}'", 1, fileName);
                                    log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Info, correlationId, Constants.AddingRecord, DateTime.UtcNow,
                                     fileName, Enums.OperationType.Add, 0);
                                    recordLogs.Add(log);
                                }
                                else
                                    continue;
                            }
                            catch (Exception ex)
                            {
                                updateCommandString = String.Format("UPDATE [Sheet1$] Set Status={0} where UploadedFileName='{1}'", 0, fileName);
                                log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Error, correlationId, "Some error ocurred while uploading a file. " + ex.Message,
                                    DateTime.UtcNow,
                                  fileName, Enums.OperationType.Add, 0);
                                recordLogs.Add(log);
                            }
                            finally
                            {
                                if (!string.IsNullOrEmpty(updateCommandString))
                                {
                                    updateCommand.CommandText = updateCommandString;
                                    updateCommand.ExecuteNonQuery();
                                }
                                count++; i++;
                                if (count == 10)
                                {
                                    Console.WriteLine(String.Format("{0} Files uploaded.", i));
                                    count = 0;
                                }
                                Helper.WriteDocumentLogsToLocalFile(recordLogs, Helper.ProjectDocumentLogFilePath);
                            }
                        }
                        connection.Close();
                    }
                }
            }
            catch (WebException we)
            {
                Console.WriteLine("Some error has occured while connecting to SharePoint Site: " + we.Message);
                log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Error, correlationId, we.Message, DateTime.UtcNow);
                Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);
            }
        }

        public void ApplyMetaDataToDocuments()
        {

            AbKLog log;
            log = Helper.ConstructLog(Enums.MigrationModule.Documents, Enums.LogType.Info, correlationId, "Applying Metadata to documents", DateTime.UtcNow);
            Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);

            Console.WriteLine("Applying Metadata to documents");
            try
            {
                string directoryPath = Helper.GetAppSettingValue(Constants.DocumentsDirectoryPath);
                string documentsListName = Helper.GetAppSettingValue(Constants.DocumentListKey);
                using (ClientContext clientContext = new ClientContext(knowledgeHubWebUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(emailId, Helper.GetPasswordFromConsoleInput(password));
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List projectList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.ProjectListKey));
                    clientContext.Load(projectList);
                    clientContext.ExecuteQuery();
                    List documentList = web.Lists.GetByTitle(Helper.GetAppSettingValue(Constants.DocumentListKey));
                    clientContext.Load(documentList);
                    clientContext.ExecuteQuery();
                    using (SqlConnection sqlConnetion = new SqlConnection(Helper.GetConnectionString(Constants.ConnectionStringKey)))
                    {
                        string queryStatement = "SELECT * FROM " + Helper.GetAppSettingValue(Constants.ProjectDocumentsViewKey)
                            + " where ProjectNumber in ('17288','19503','20152','19912','18210','16488','07544','17916','18274','19410','16590','20330','20149','20479','13243')";
                        using (SqlCommand command = new SqlCommand(queryStatement, sqlConnetion))
                        {
                            sqlConnetion.Open();
                            SqlDataReader reader = command.ExecuteReader();
                            int count = 0; int i = 0; string lastProjectNumber = string.Empty;
                            ListItem project = null;
                            Enums.OperationType typeOfOperation = Enums.OperationType.ApplyingMetadata;
                            List<AbKLog> recordLogs;
                            while (reader.Read())
                            {
                                recordLogs = new List<AbKLog>();
                                string projectNumber = string.Empty;
                                string fileName = string.Empty;
                                ListItem document = null;
                                string termValueString = string.Empty;
                                string notResolvedUsers = string.Empty;

                                try
                                {
                                    projectNumber = Helper.SafeGetString(reader, "ProjectNumber");
                                    fileName = Helper.SafeGetString(reader, "UploadedFileName");
                                    if (string.IsNullOrEmpty(projectNumber) || string.IsNullOrEmpty(fileName))
                                        continue;

                                    string documentAuthor = Helper.SafeGetString(reader, "Author");
                                    string documentTitle = Helper.SafeGetString(reader, "Title");
                                    // List<FieldUserValue> Author = GetSPUserCollection(clientContext, web, documentAuthor, out notResolvedUsers);

                                    string fileCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='LinkFilename' /><Value Type='Text'>" + fileName + @"</Value></Eq></Where>
                                                              </Query><ViewFields><FieldRef Name='ProjectNumber' /><FieldRef Name='KHProject' /><FieldRef Name='Title' />"
                                                            + "<FieldRef Name='DocumentDate' /><FieldRef Name='KHAbtkAuthor' /><FieldRef Name='KHFormerAbrAuthor' />"
                                                            + "<FieldRef Name='globalAbtCapabilities' /><FieldRef Name='globalAbtOrganization' /><FieldRef Name='globalClientTypes' /><FieldRef Name='globalGeographicLocations' />"
                                                            + "<FieldRef Name='globalProjectDemographics' /><FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
                                    ListItemCollection documentListItemCollection = GetItems(clientContext, documentList, fileCamlQuery);
                                    if (documentListItemCollection != null && documentListItemCollection.Count > 0)
                                        document = documentListItemCollection[0];
                                    else
                                    {
                                        log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Error, correlationId, "Document not found.",
                                            DateTime.UtcNow, fileName, typeOfOperation, 0);
                                        recordLogs.Add(log);
                                        continue;
                                    }

                                    if (lastProjectNumber != projectNumber)
                                    {
                                        string projectCamlQuery = @"<View><Query><Where><Eq><FieldRef Name='ProjectOracleNumber' /><Value Type='Text'>" + projectNumber + @"</Value></Eq></Where>
                                                              </Query><ViewFields><FieldRef Name='ID' /><FieldRef Name = 'globalAbtCapabilities' /><FieldRef Name = 'globalAbtOrganization' />"
                                                              + "<FieldRef Name='globalClientTypes' /><FieldRef Name = 'globalGeographicLocations' /><FieldRef Name='globalProjectDemographics' />"
                                                              + "<FieldRef Name='globalSubjectMatterAreas' /></ViewFields></View>";
                                        ListItemCollection projectListItemCollection = GetItems(clientContext, projectList, projectCamlQuery);
                                        if (projectListItemCollection != null && projectListItemCollection.Count > 0)
                                            project = projectListItemCollection[0];
                                        else
                                        {
                                            log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Error, correlationId, "Project not found.",
                                                DateTime.UtcNow, fileName, typeOfOperation, 0);
                                            recordLogs.Add(log);
                                            continue;
                                        }
                                    }
                                    lastProjectNumber = projectNumber;
                                    FieldLookupValue projectLookup = new FieldLookupValue();
                                    projectLookup.LookupId = project.Id;
                                    document["ProjectNumber"] = projectLookup;
                                    document["KHProject"] = projectLookup;
                                    document["DocumentDate"] = Helper.SafeGetDateTime(reader, "FileDate");
                                    document["Title"] = String.IsNullOrEmpty(documentTitle) ? "" : StringExt.Truncate(documentTitle, 255);

                                    //if (Author != null && Author.Count != 0)
                                    //{
                                    //    document["KHAbtkAuthor"] = Author.ToArray();//Helper.SafeGetString(reader, "ProjectDirector");
                                    //}
                                    document["KHFormerAbrAuthor"] = documentAuthor;

                                    TaxonomyFieldValueCollection currentCollection = project["globalAbtCapabilities"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalAbtCapabilities");
                                    currentCollection = project["globalAbtOrganization"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalAbtOrganization");
                                    currentCollection = project["globalClientTypes"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalClientTypes");
                                    currentCollection = project["globalGeographicLocations"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalGeographicLocations");
                                    currentCollection = project["globalProjectDemographics"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalProjectDemographics");
                                    currentCollection = project["globalSubjectMatterAreas"] as TaxonomyFieldValueCollection;
                                    termValueString = GetTaxonomyStringFromCollection(currentCollection);
                                    if (!string.IsNullOrEmpty(termValueString))
                                        SetTaxonomyCollection(clientContext, documentList, termValueString, document, "globalSubjectMatterAreas");

                                    document.Update();
                                    log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Info, correlationId, Constants.RecordAddedUpdated, DateTime.UtcNow,
                                     fileName, typeOfOperation, document.Id);
                                    recordLogs.Add(log);
                                }

                                catch (Exception e)
                                {
                                    log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Error, correlationId, Constants.ErrorRecordAddedUpdated + " " + e.Message,
                                        DateTime.UtcNow, fileName, typeOfOperation, (document == null || document.ServerObjectIsNull == null || !document.ServerObjectIsNull.Value) ? 0 : document.Id);
                                    recordLogs.Add(log);
                                }
                                finally
                                {
                                    count++; i++;
                                    if (count == 10)
                                    {
                                        Console.WriteLine(String.Format("Metadata applied to {0} documents.", i));
                                        count = 0;
                                    }
                                    Helper.WriteDocumentLogsToLocalFile(recordLogs, Helper.ProjectDocumentLogFilePath);
                                }
                            }
                            sqlConnetion.Close();
                            log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Info, correlationId, "Project import is complete.", DateTime.UtcNow);
                            Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);
                            Console.WriteLine("Metadata application finished");
                        }
                    }
                }
            }
            catch (WebException we)
            {
                Console.WriteLine("Some error has occured while connecting to SharePoint Site: " + we.Message);
                log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Error, correlationId, we.Message, DateTime.UtcNow);
                Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);
            }
            catch (Exception e)
            {
                log = Helper.ConstructLog(Enums.MigrationModule.Metadata, Enums.LogType.Error, correlationId, e.Message, DateTime.UtcNow);
                Helper.WriteDocumentLogToLocalFile(log, Helper.ProjectDocumentLogFilePath);
            }
        }

        public ListItemCollection GetItems(ClientContext clientContext, List list, string query)
        {
            CamlQuery projectCamlQuery = new CamlQuery();
            projectCamlQuery.ViewXml = query;
            ListItemCollection projectListItemCollection = list.GetItems(projectCamlQuery);
            clientContext.Load(projectListItemCollection);
            clientContext.ExecuteQuery();
            return projectListItemCollection;
        }

        public User GetSPUser(ClientContext clientContext, Web web, string emailId)
        {
            User user = null;
            try
            {
                if (!string.IsNullOrEmpty(emailId))
                {
                    user = clientContext.Web.EnsureUser(Constants.membership + emailId);
                    clientContext.Load(user);
                    clientContext.ExecuteQuery();
                }
            }
            catch
            { }
            return user;
            //User user = null;
            //try
            //{
            //    if (!string.IsNullOrEmpty(emailId))
            //    {
            //        var userPrincipal = Microsoft.SharePoint.Client.Utilities.Utility.ResolvePrincipal(clientContext,
            //               web,
            //             emailId, // normal login name
            //               Microsoft.SharePoint.Client.Utilities.PrincipalType.User,
            //               Microsoft.SharePoint.Client.Utilities.PrincipalSource.All,
            //               users,
            //               false);

            //        clientContext.ExecuteQuery();
            //        if (userPrincipal == null || userPrincipal.Value == null || string.IsNullOrEmpty(userPrincipal.Value.LoginName))
            //            return user;
            //        // get a User instance based on the encoded login name from userPrincipal
            //        user = users.GetByLoginName(userPrincipal.Value.LoginName);
            //        clientContext.Load(user);
            //        clientContext.ExecuteQuery();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    string a = "0";
            //}
            //return user;
        }
        public List<FieldUserValue> GetSPUserCollection(ClientContext clientContext, Web web, string emailIds, out string notResolvedUsers)
        {
            List<FieldUserValue> users = null; notResolvedUsers = "";
            try
            {
                if (!string.IsNullOrEmpty(emailIds))
                {
                    string[] emails = emailIds.Split(';');
                    users = new List<FieldUserValue>();

                    foreach (string email in emails)
                    {
                        var user = GetSPUser(clientContext, web, email);

                        if (user != null && user.ServerObjectIsNull != null && !user.ServerObjectIsNull.Value)
                        {
                            var userValue = new FieldUserValue();
                            userValue.LookupId = user.Id;
                            users.Add(userValue);
                        }
                        else
                        {
                            notResolvedUsers += email;
                        }

                    }
                }
            }

            catch (Exception ex)
            {
                string a = "1";
            }
            return users;
        }
        public string GetTaxonomyStringFromCollection(TaxonomyFieldValueCollection collection)
        {
            string termValueString = string.Empty;
            if (collection == null || collection.Count == 0)
                return termValueString;
            foreach (TaxonomyFieldValue tv in collection)
            {
                termValueString += tv.WssId + ";#" + tv.Label + "|" + tv.TermGuid + ";#";
            }
            termValueString = termValueString.TrimEnd('#').TrimEnd(';');
            return termValueString;
        }

        public void SetTaxonomyCollection(ClientContext clientContext, List projectDescriptionList, string termValueString, ListItem descriptionItem, string columnName)
        {
            Field field = projectDescriptionList.Fields.GetByInternalNameOrTitle(columnName);
            TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);
            TaxonomyFieldValueCollection currentCollection = new TaxonomyFieldValueCollection(clientContext, termValueString, txField);
            txField.SetFieldValueByValueCollection(descriptionItem, currentCollection);
        }

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

        public static Microsoft.SharePoint.Client.File UploadFileSlicePerSlice(ClientContext ctx, string libraryName, string fileName, int fileChunkSizeInMB = 10)
        {
            // Each sliced upload requires a unique id
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file
            string uniqueFileName = Path.GetFileName(fileName);


            // Get to folder to upload into 
            List docs = ctx.Web.Lists.GetByTitle(libraryName);
            ctx.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // File object 
            Microsoft.SharePoint.Client.File uploadFile;

            // Calculate block size in bytes
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            // Get the information about the folder that will hold the file
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();


            // Get the size of the file
            long fileSize = new FileInfo(fileName).Length;

            if (fileSize <= blockSize)
            {
                // Use regular approach
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = uniqueFileName;
                    fileInfo.Overwrite = true;
                    uploadFile = docs.RootFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // return the file object for the uploaded file
                    return uploadFile;
                }
            }
            else
            {
                // Use large file upload approach
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from filesystem in blocks 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // We've reached the end of the file
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = uniqueFileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = docs.RootFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // we can only start the upload once
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to our file
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(docs.RootFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // return the file object for the uploaded file
                                        return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // update fileoffset for the next slice
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }

                }
            }

            return null;
        }
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        //public static void ProcessDirectory(string targetDirectory, ClientContext clientContext, string documentLibraryName)
        //{
        //    // Process the list of files found in the directory.
        //    string[] fileEntries = Directory.GetFiles(targetDirectory);
        //    foreach (string fileName in fileEntries)
        //        ProcessFile(fileName,clientContext, documentLibraryName);

        //    // Recurse into subdirectories of this directory.
        //    string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
        //    foreach (string subdirectory in subdirectoryEntries)
        //        ProcessDirectory(subdirectory);
        //}

        //// Insert logic for processing found files here.
        //public static void ProcessFile(string path)
        //{
        //    Console.WriteLine("Processed file '{0}'.", path);
        //}

    }
}
