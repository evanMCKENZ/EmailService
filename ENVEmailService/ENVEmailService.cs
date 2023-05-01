using System;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;
using System.Data.SqlClient;
using System.IO;

namespace ENVEmailService
{
    //Service State constants
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]

    //Service status variables
    public struct ServiceStatus
    {
        public int dwServiceType;
        public ServiceState dwCurrentState;
        public int dwControlsAccepted;
        public int dwWin32ExitCode;
        public int dwServiceSpecificExitCode;
        public int dwCheckPoint;
        public int dwWaitHint;
    };

    public partial class ENVEmailService : ServiceBase
    {

        readonly EventLog eventLog1;

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        public ENVEmailService()
        {
            InitializeComponent();
            eventLog1 = new EventLog();                     //create event log and link all inputs
            if (!EventLog.SourceExists("ENVSource"))
            {
                EventLog.CreateEventSource(
                    "ENVSource", "EMEmailLog");         //create Event Viewer object
            }
            eventLog1.Source = "ENVSource";
            eventLog1.Log = "EMEmailLog";               //set variables for event log
        }

        protected override void OnStart(string[] args)
        {
            //Set service state to pending start
            ServiceStatus serviceStatus = new ServiceStatus
            {
                dwCurrentState = ServiceState.SERVICE_START_PENDING,
                dwWaitHint = 100000
            };
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry("In OnStart.");

            EmailWeekly();

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");             //write to event log with current position
        }

        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");             //write to event log with current position
        }

        protected override void OnShutdown()
        {
            eventLog1.WriteEntry("In OnShutdown.");                 //write to event log with current position
        }

        protected override void OnPause()
        {
            eventLog1.WriteEntry("In OnPause.");                        //write to event log with current position
        }

        public void EmailWeekly()
        {
            // ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- //

            /* Connecting to SQL Servers and Retrieving Data
            *         - Create connection to both servers and databases
            *         - Execute created SQL statements
            *         - Confirm data recovery
            *         - Close connections and functions
            */

            //connection strings for IBA SQL Server
            string IBAconnString1 = @"Data Source = XXXXXXXXXXXXX; Initial Catalog = NK_12Hr; MultipleActiveResultSets = True; Integrated Security = False; Persist Security Info = True; User ID = XXXXXXXXXXXXX; Password = XXXXXXXXXXXXX;";
            string IBAconnString2 = @"Data Source = XXXXXXXXXXXXX; Initial Catalog = NK_PressDrop; MultipleActiveResultSets = True; Integrated Security = False; Persist Security Info = True; User ID = XXXXXXXXXXXXX; Password = XXXXXXXXXXXXX;";

            SqlConnection ibaConn1 = new SqlConnection(IBAconnString1);             //SQLConnection
            SqlConnection ibaConn2 = new SqlConnection(IBAconnString2);             //SQLConnection

            //connection string for ENV SQL Server
            string ENVconnString = @"Data Source = XXXXXXXXXXXXX; Initial Catalog = ENV; MultipleActiveResultSets = True; Integrated Security = False; Persist Security Info = False; User ID = XXXXXXXXXXXXX; Password = XXXXXXXXXXXXX;";

            SqlConnection envConn = new SqlConnection(ENVconnString);             //SQLConnection

            string path = @"C:\errEmailPage.txt";

            try
            {
                using (ibaConn1)
                using(ibaConn2)
                using(envConn)
                {
                    ibaConn1.Open();                     //open connection to IBA database
                    ibaConn2.Open();
                    envConn.Open();                     //open connection to ENV database

                    string ibaSQLstring1 = "DECLARE @date DATETIME, @weekly DATETIME set @date = GETDATE() set @weekly = DATEADD(day, -7, CAST(GETDATE() AS date)) select _TimeStamp as Timestamp, BH1_North_FA, BH1_South_FA, BH2_East_FA, BH2_West_FA, BH1_CC_Damper, BH1_SC_Damper, BH2_NC_Damper from ibaFile where _TimeStamp >= @weekly AND _TimeStamp <= @date order by Timestamp DESC";
                    string ibaSQLstring2 = "DECLARE @date DATETIME, @weekly DATETIME set @date = GETDATE() set @weekly = DATEADD(day, -7, CAST(GETDATE() AS date)) select[DateTime] as Timestamp, Baghouse_1 as BH1_15Avg, Baghouse_2 as BH2_15Avg from NK_PressDrop.dbo.PressureDrop_15minAvg where [DateTime] >= @weekly AND [DateTime] <= @date order by Timestamp DESC";
                    string envSQLstring1 = "DECLARE @date DATETIME, @weekly DATETIME set @date = GETDATE() set @weekly = DATEADD(day, -7, CAST(GETDATE() AS date)) select[DateTime] as Timestamp, North_Canopy, Center_Canopy, South_Canopy from ENV.dbo.Canopy_Damper where [DateTime] >= @weekly AND [DateTime] <= @date order by [DateTime] DESC";
                    string envSQLstring2 = "DECLARE @date DATETIME, @weekly DATETIME set @date = GETDATE() set @weekly = DATEADD(day, -7, CAST(GETDATE() AS date)) select[TimeStamp] as Timestamp, BH1_North, BH1_South, BH2_East, BH2_West from ENV.dbo.Fan_Amps where [TimeStamp] >= @weekly AND [TimeStamp] <= @date order by [TimeStamp] DESC";
                    string envSQLstring3 = "DECLARE @date DATETIME, @weekly DATETIME set @date = GETDATE() set @weekly = DATEADD(day, -7, CAST(GETDATE() AS date)) select[DateTime] as Timestamp, [Baghouse 1 15m] as BH1_15m, [Baghouse 2 15m] as BH2_15m, [Baghouse 1 1h] as BH1_1h, [Baghouse 2 1h] as BH2_1h from ENV.dbo.Pressure_Drop where [DateTime] >= @weekly AND [DateTime] <= @date order by [DateTime] DESC";

                    SqlCommand ibaCom1 = new SqlCommand(ibaSQLstring1)
                    {
                        Connection = ibaConn1                          //link connection to command
                    };
                    SqlCommand ibaCom2 = new SqlCommand(ibaSQLstring2)
                    {
                        Connection = ibaConn2                          //link connection to command
                    };
                    SqlCommand envCom1 = new SqlCommand(envSQLstring1)
                    {
                        Connection = envConn                          //link connection to command
                    };
                    SqlCommand envCom2 = new SqlCommand(envSQLstring2)
                    {
                        Connection = envConn                          //link connection to command
                    };
                    SqlCommand envCom3 = new SqlCommand(envSQLstring3)
                    {
                        Connection = envConn                          //link connection to command
                    };

                    SqlDataReader ibaReader1 = ibaCom1.ExecuteReader();             //read from sql database and store results in reader array
                    SqlDataReader ibaReader2 = ibaCom2.ExecuteReader();             //read from sql database and store results in reader array
                    SqlDataReader envReader1 = envCom1.ExecuteReader();             //read from sql database and store results in reader array
                    SqlDataReader envReader2 = envCom2.ExecuteReader();             //read from sql database and store results in reader array
                    SqlDataReader envReader3 = envCom3.ExecuteReader();             //read from sql database and store results in reader array

                    DataTable ibaTable1 = new DataTable();
                    ibaTable1.Columns.Clear();
                    DataTable ibaTable2 = new DataTable();
                    ibaTable2.Columns.Clear();
                    DataTable envTable1 = new DataTable();                      //create DataDables for each SQL data query
                    envTable1.Columns.Clear();                                  //confirm all tables are clear for data
                    DataTable envTable2 = new DataTable();
                    envTable2.Columns.Clear();
                    DataTable envTable3 = new DataTable();
                    envTable3.Columns.Clear();

                    ibaTable1.Load(ibaReader1);
                    ibaTable2.Load(ibaReader2);
                    envTable1.Load(envReader1);                         //load DataTables with SQLDataReader results   
                    envTable2.Load(envReader2);
                    envTable3.Load(envReader3);

                    if (ibaTable1.Rows.Count>0 && ibaTable2.Rows.Count>0 && envTable1.Rows.Count>0 && envTable2.Rows.Count>0 && envTable3.Rows.Count>0)
                    {

            // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- //

            /* Formatting Data from SQL and Sending Emails
            *         - Create List of database entries from both tables
            *         - Create SMTP client for email sending
            *         - Format data from SQL Reader
            *         - Send emails and close all links
            */

                        string todayDate = DateTime.Now.ToString("M/d/yyyy");                //will return current date only, no time

                        MailAddress fromMail = new MailAddress("XXXXXXXXXXXXXXXXXXXXX");          //create new email address to send from
                        MailAddress toMail = new MailAddress("XXXXXXXXXXXXXXXXXXXXX");          //create new email address to send from
                        string subject = "Environmental Data for the Week Ending " + todayDate;
                        string body = "See data broken down by database <br><br>";
                        body += "<h3>IBA Table 1 - 12Hr Averages</h3>";
                        body += CreateHtmlTableFormat(ibaTable1);
                        body += "<br>";
                        body += "<h3>IBA Table 2 - Pressure Drop</h3>";
                        body += CreateHtmlTableFormat(ibaTable2);
                        body += "<br>";
                        body += "<h3>ENV Table 1 - Canopy Damper</h3>";             //HTML formatting 
                        body += CreateHtmlTableFormat(envTable1);               //pass DataTable to method to produce String object 
                        body += "<br>";
                        body += "<h3>ENV Table 2 - Fan Amps</h3>";
                        body += CreateHtmlTableFormat(envTable2);
                        body += "<br>";
                        body += "<h3>ENV Table 3 - Pressure Drop</h3>";
                        body += CreateHtmlTableFormat(envTable3);
                        body += "<br>";
                        body += "<br>";
                        body += "<h3> This E-Mail is automatically sent. Please do not reply. </h3>";

                        MailMessage mailMessage = new MailMessage(fromMail, toMail);
                        mailMessage.Subject = subject;
                        mailMessage.Body = body;
                        mailMessage.IsBodyHtml = true;

                            
                        SmtpClient smtpClient = new SmtpClient("smtp-mail.outlook.com", 587)            //FIND NUCOR SMTP SERVER INFO
                        {
                            EnableSsl = true,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            UseDefaultCredentials = false,
                            Credentials = new NetworkCredential("XXXXXXXXXXXX", "XXXXXXXXXXXXXX")       //admin credentials would be good here
                        };
                        try
                        {
                            smtpClient.Send(mailMessage);
                        }
                        catch (Exception ex)
                        {
                            eventLog1.WriteEntry("Error: " + ex.Message);
                            File.WriteAllText(path, ex.Message);
                            File.AppendAllText(path, ex.StackTrace);

                        }
                    }
                }

            }
            catch (Exception e)             //catchall for any errors or exceptions
            {
                eventLog1.WriteEntry("Error: " + e.Message);           //print stacktrace to event log
                File.WriteAllText(path, e.Message);
                File.AppendAllText(path, e.StackTrace);
            }
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- //

        /* Email Data Formatting for Readability
        *         - Create HTML outlines and table
        *         - Add column headers from SQL table
        *         - Format rows from SQL Reader into table cells
        *         - Return a StringBuilder object with inline HTML
        */
        public string CreateHtmlTableFormat(DataTable dt)
        {
            string tab = "\t";

            StringBuilder sb = new StringBuilder();                 //return object with inline formatting

            sb.AppendLine("<html>");
            sb.AppendLine(tab + "<body>");
            sb.AppendLine(tab + tab + "<table>");                   //HTML syntax settings
            sb.AppendLine(tab + tab + "<style> tr td:first-child {padding-left:0px;} td { padding: 00px 0px 0px 20px; } </style>");     //INLINE CSS for column spacing

            // headers.
            sb.Append(tab + tab + tab + "<tr>");

            foreach (DataColumn dc in dt.Columns)
            {
                sb.AppendFormat("<td>   {0}   </td>", dc.ColumnName);                 //add column header titles
            }

            sb.AppendLine("</tr>");

            // data rows
            foreach (DataRow dr in dt.Rows)
            {
                sb.Append(tab + tab + tab + "<tr>");

                foreach (DataColumn dc in dt.Columns)
                {
                    string cellValue = dr[dc] != null ? dr[dc].ToString() : "   ";             //add each cell of SQL table to HTML cell
                    sb.AppendFormat("<td><center>  {0}  </center></td>", cellValue);            //center all data entries
                }

                sb.AppendLine("</tr>");
            }

            sb.AppendLine(tab + tab + "</table>");
            sb.AppendLine(tab + "</body>");
            sb.AppendLine("</html>");                   //close HTML syntax

            return sb.ToString();
        }
    }
}
