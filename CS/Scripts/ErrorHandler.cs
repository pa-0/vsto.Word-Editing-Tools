using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
//using Excel = Microsoft.Office.Interop.Excel;
using log4net;
using log4net.Config;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

// <summary> 
// This namespaces if for generic application classes
// </summary>
namespace EditTools.Scripts
{
    /// <summary> 
    /// Used to handle exceptions
    /// </summary>
    public class ErrorHandler
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorHandler));

        /// <summary>
        /// Applies a new path for the log file by FileAppender name
        /// </summary>
        public static void SetLogPath()
        {
            XmlConfigurator.Configure();
            log4net.Repository.Hierarchy.Hierarchy h = (log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository();
            //string logFileName = System.IO.Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".log");
            string logFileName = System.IO.Path.Combine(@"C:\Temp\", AssemblyInfo.Product + ".log");
            foreach (var a in h.Root.Appenders)
            {
                if (a is log4net.Appender.FileAppender)
                {
                    if (a.Name.Equals("FileAppender"))
                    {
                        log4net.Appender.FileAppender fa = (log4net.Appender.FileAppender)a;
                        fa.File = logFileName;
                        fa.ActivateOptions();
                    }
                }
            }
        }

        /// <summary>
        /// Create a log record to track which methods are being used.
        /// </summary>
        public static void CreateLogRecord()
        {
            try
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(1);
                System.Reflection.MethodBase caller = sf.GetMethod();
                string currentProcedure = (caller.Name).Trim();
                log.Info("[PROCEDURE]=|" + currentProcedure + "|[USER NAME]=|" + Environment.UserName + "|[MACHINE NAME]=|" + Environment.MachineName);
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary> 
        /// Used to produce an error message and create a log record
        /// <example>
        /// <code lang="C#">
        /// ErrorHandler.DisplayMessage(ex);
        /// </code>
        /// </example> 
        /// </summary>
        /// <param name="ex">Represents errors that occur during application execution.</param>
        /// <param name="isSilent">Used to show a message to the user and log an error record or just log a record.</param>
        /// <remarks></remarks>
        public static void DisplayMessage(Exception ex, Boolean isSilent = false)
        {
            System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(1);
            System.Reflection.MethodBase caller = sf.GetMethod();
            string currentProcedure = (caller.Name).Trim();
            string currentFileName = AssemblyInfo.GetCurrentFileName();
            string errorMessageDescription = ex.ToString();
            errorMessageDescription = System.Text.RegularExpressions.Regex.Replace(errorMessageDescription, @"\r\n+", " "); //the carriage returns were messing up my log file
            string msg = "Contact your system administrator. A record has been created in the log file." + Environment.NewLine;
            msg += "Procedure: " + currentProcedure + Environment.NewLine;
            msg += "Description: " + ex.ToString() + Environment.NewLine;
            log.Error("[PROCEDURE]=|" + currentProcedure + "|[USER NAME]=|" + Environment.UserName + "|[MACHINE NAME]=|" + Environment.MachineName + "|[FILE NAME]=|" + currentFileName + "|[DESCRIPTION]=|" + errorMessageDescription);
            if (isSilent == false)
            {
                MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary> 
        /// Check if the object is a date     
        /// </summary>
        /// <param name="expression">Represents the cell value </param>
        /// <returns>A method that returns true or false if there is a valid date </returns> 
        /// <remarks></remarks>
        public static bool IsDate(object expression)
        {
            if (expression != null)
            {
                if (expression is DateTime)
                {
                    return true;
                }
                if (expression is string)
                {
                    DateTime time1;
                    return DateTime.TryParse((string)expression, out time1);
                }
            }
            return false;
        }

        /// <summary> 
        /// Check if the object is a time     
        /// </summary>
        /// <param name="expression">Represents the cell value </param>
        /// <returns>A method that returns true or false if there is a valid time </returns> 
        /// <remarks></remarks>
        public static bool IsTime(object expression)
        {
            try
            {
                string timeValue = Convert.ToString(expression);
                //timeValue = String.Format("{0:" + Properties.Settings.Default.Table_ColumnFormatTime + "}", expression);
                //timeValue = timeValue.ToString(Properties.Settings.Default.Table_ColumnFormatTime, System.Globalization.CultureInfo.InvariantCulture);
                //timeValue = timeValue.ToString(Properties.Settings.Default.Table_ColumnFormatTime);
                //string timeValue = expression.ToString().Format(Properties.Settings.Default.Table_ColumnFormatTime);
                TimeSpan time1;
                //return TimeSpan.TryParse((string)expression, out time1);
                return TimeSpan.TryParse(timeValue, out time1);
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}