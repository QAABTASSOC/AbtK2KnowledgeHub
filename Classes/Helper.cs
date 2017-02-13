using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AbtK2KnowledgeHub_OneTime.Classes
{
    public class Helper
    {
        public static string GeneralLogFilePath
        {
            get
            {
                // return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAppSettingValue(Constants.LogFilePathKey));
                return @"C:\Deepa\Prakash.Abt.Migration\AbtK2KnowledgeHub-OneTime\Log\log.txt";
            }
        }
        public static string ProjectLogFilePath
        {
            get
            {
                // return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAppSettingValue(Constants.LogFilePathKey));
                return @"C:\Deepa\Prakash.Abt.Migration\AbtK2KnowledgeHub-OneTime\Log\ProjectLogs.txt";
            }
        }
        public static string ProjectStaffLogFilePath
        {
            get
            {
                // return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAppSettingValue(Constants.LogFilePathKey));
                return @"C:\Deepa\Prakash.Abt.Migration\AbtK2KnowledgeHub-OneTime\Log\ProjectStaffLogs.txt";
            }
        }
        public static string ProjectDescriptionLogFilePath
        {
            get
            {
                // return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAppSettingValue(Constants.LogFilePathKey));
                return @"C:\Deepa\Prakash.Abt.Migration\AbtK2KnowledgeHub-OneTime\Log\ProjectDescriptionLogs.txt";
            }
        }
        public static string ProjectDocumentLogFilePath
        {
            get
            {
                // return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GetAppSettingValue(Constants.LogFilePathKey));
                return @"C:\Deepa\Prakash.Abt.Migration\AbtK2KnowledgeHub-OneTime\Log\ProjectDocumentLogs.txt";
            }
        }
        /// <summary>
        /// Write logs to Localfile
        /// </summary>
        /// <param name="logs"></param>
        /// <param name="logFilePath"></param>
        /// <returns></returns>
        public static bool WriteLogsToLocalFile(List<AbKLog> logs, string filePath)
        {
            if (logs == null)
                return true;

            try
            {
                foreach (AbKLog log in logs)
                {
                    WriteLogToLocalFile(log, filePath);
                }
            }
            catch (Exception) { return false; }

            return true;
        }
        public static bool WriteLogToLocalFile(AbKLog log, string filePath)
        {
            if (log == null)
                return true;
            try
            {
                string strlog = Environment.NewLine + Convert.ToString(log.Type) + "|" + log.CreatedOn + "|" +
                    log.CorelationId + "," + Convert.ToString(log.Category) + "|" + log.Message + "|" + log.EntityID + "|" + log.OperationType + "|" + log.SharePointID;
                if (File.Exists(filePath))
                    File.AppendAllText(filePath, strlog);
            }
            catch (Exception) { return false; }

            return true;
        }

        /// <summary>
        /// Write logs to Localfile
        /// </summary>
        /// <param name="logs"></param>
        /// <param name="logFilePath"></param>
        /// <returns></returns>
        public static bool WriteDocumentLogsToLocalFile(List<AbKLog> logs, string logFilePath)
        {
            if (logs == null)
                return true;
            try
            {
                foreach (var log in logs)
                {
                    WriteDocumentLogToLocalFile(log, logFilePath);
                }
            }
            catch (Exception) { return false; }

            return true;
        }
        /// <summary>
        /// Write logs to Localfile
        /// </summary>
        /// <param name="logs"></param>
        /// <param name="logFilePath"></param>
        /// <returns></returns>
        public static bool WriteDocumentLogToLocalFile(AbKLog log, string logFilePath)
        {
            if (log == null)
                return true;
            try
            {
                    string strlog = Environment.NewLine + Convert.ToString(log.Type) + "|" + log.CreatedOn + "|" +
                        log.CorelationId + "|" + Convert.ToString(log.Category) + "|" + log.Message + "|" + log.EntityName + "|" + log.OperationType + "|" + log.SharePointID;
                    if (File.Exists(logFilePath))
                        File.AppendAllText(logFilePath, strlog);
            }
            catch (Exception) { return false; }

            return true;
        }
        public static AbKLog ConstructLog(Enums.MigrationModule Category, Enums.LogType Type, Guid CorelationId,
          string Message, DateTime CreatedOn)
        {
            var log = new AbKLog();
            log.Category = Category;
            log.Type = Type;
            log.CorelationId = CorelationId;
            log.Message = Message;
            log.CreatedOn = CreatedOn;
            return log;
        }
        public static AbKLog ConstructLog(Enums.MigrationModule Category, Enums.LogType Type, Guid CorelationId,
         string Message, DateTime CreatedOn, long EntityID, Enums.OperationType OperationType, int SharePointID)
        {
            var log = ConstructLog(Category, Type, CorelationId, Message, CreatedOn);
            log.EntityID = EntityID;
            log.OperationType = OperationType;
            log.SharePointID = SharePointID;
            return log;
        }
        public static AbKLog ConstructLog(Enums.MigrationModule Category, Enums.LogType Type, Guid CorelationId,
      string Message, DateTime CreatedOn, string EntityName, Enums.OperationType OperationType, int SharePointID)
        {
            var log = ConstructLog(Category, Type, CorelationId, Message, CreatedOn);
            log.EntityName = EntityName;
            log.OperationType = OperationType;
            log.SharePointID = SharePointID;
            return log;
        }
        /// <summary>
        /// Get app setting value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetConnectionString(string key)
        {
            string value = string.Empty;
            try
            {
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "AbtK2KnowledgeHub-OneTime.exe"));
                ConnectionStringSettingsCollection connectionStringCollection = config.ConnectionStrings.ConnectionStrings;
                value = connectionStringCollection[key].ConnectionString;
                //value = ConfigurationManager.AppSettings[key];
            }
            catch
            {
                throw;
            }
            return value;
        }


        /// <summary>
        /// Get app setting value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetAppSettingValue(string key)
        {
            string value = string.Empty;
            try
            {
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "AbtK2KnowledgeHub-OneTime.exe"));
                KeyValueConfigurationCollection svcSettings = config.AppSettings.Settings;
                value = svcSettings[key].Value;
                //value = ConfigurationManager.AppSettings[key];
            }
            catch { }
            return value;
        }
        public static int? SafeGetInt32(SqlDataReader reader,
                                  string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetInt32(ordinal);
            }
            return null;
        }
        public static Int64? SafeGetInt64(SqlDataReader reader,
                                 string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetInt64(ordinal);
            }
            return null;
        }
        public static string SafeGetString(SqlDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetString(ordinal);
            }
            return null;
        }
        public static DateTime? SafeGetDateTime(SqlDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetDateTime(ordinal);
            }
            return null;
        }
        public static decimal? SafeGetDecimal(SqlDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetDecimal(ordinal);
            }
            return null;
        }
        public static bool? SafeGetBool(SqlDataReader reader, string columnName)
        {
            int ordinal = reader.GetOrdinal(columnName);

            if (!reader.IsDBNull(ordinal))
            {
                return reader.GetBoolean(ordinal);
            }
            return null;
        }
        public static SecureString GetPasswordFromConsoleInput(string password)
        {
            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            char[] securePasswordArray = password.ToCharArray();
            for (int i = 0; i < securePasswordArray.Length; i++)
            {
                securePassword.AppendChar(securePasswordArray[i]);
            }
            return securePassword;
        }

    }
    public static class StringExt
    {
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }

}
