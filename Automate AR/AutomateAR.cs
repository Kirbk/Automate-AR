using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.ServiceProcess;
using System.Timers;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Automate_AR
{
    public partial class AutomateAR : ServiceBase
    {
        public static EventLog eventLog;
        public static int eventId = 0;
        private Mailer mailer;

        private bool dryRun = false;
        private int reminder_time = 30;
        private int overdue_time = 45;

        private List<string> allowedAddress = new List<string>();
        private List<string> ccAddress = new List<string>();

        public AutomateAR()
        {
            InitializeComponent();
            eventLog = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("AutomateAR"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "AutomateAR", "INFO");
            }
            eventLog.Source = "AutomateAR";
            eventLog.Log = "INFO";

            mailer = new Mailer();

            System.IO.Directory.CreateDirectory(@"C:\AutomateAR");

            string path = @"C:\AutomateAR\config";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("dry: true\n\nreminder: 30\nlate: 45\n\n[CC]\nckirby@augustmack.com\n\n[Allow]\nckirby@augustmack.com");
                }
            }

            List<string> config = System.IO.File.ReadAllLines(path).ToList<string>();
            char[] separators = new char[] { ' ', ':' };

            bool allowSwitch = false;
            bool ccSwitch = false;

            foreach (string line in config)
            {
                if (allowSwitch)
                {
                    if (line.Equals("[CC]", StringComparison.OrdinalIgnoreCase))
                    {
                        ccSwitch = true;
                        allowSwitch = false;
                    }
                    else
                    {
                        eventLog.WriteEntry(line, EventLogEntryType.Warning, eventId++);
                        allowedAddress.Add(line);
                    }
                }
                else if (ccSwitch)
                {

                    if (line.Equals("[Allow]", StringComparison.OrdinalIgnoreCase))
                    {
                        ccSwitch = false;
                        allowSwitch = true;
                    }
                    else
                    {
                        if (!line.Equals(""))
                        {
                            eventLog.WriteEntry("Registering CC " + line, EventLogEntryType.Information, eventId++);
                            ccAddress.Add(line);
                        }
                    }
                }

                else if (!dryRun && line.Equals("[Allow]", StringComparison.OrdinalIgnoreCase))
                {
                    ccSwitch = false;
                    allowSwitch = true;
                }
                else if (line.Equals("[CC]", StringComparison.OrdinalIgnoreCase))
                {
                    allowSwitch = false;
                    ccSwitch = true;
                }
                else
                {
                    string[] keyvalue = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                    if (keyvalue.Length == 2)
                    {
                        switch (keyvalue[0])
                        {
                            case "dry":
                                dryRun = Convert.ToBoolean(keyvalue[1]);
                                break;
                            case "reminder":
                                reminder_time = Convert.ToInt32(keyvalue[1]);
                                break;
                            case "late":
                                overdue_time = Convert.ToInt32(keyvalue[1]);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }

        protected override void OnStart(string[] args)
        {
            // Set up a timer that triggers every day.
            Timer timer = new Timer();
            timer.Interval = 86_400_000; // 1 Day
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();

            foreach (string cc in ccAddress)
                eventLog.WriteEntry("CCing " + cc, EventLogEntryType.Information, eventId++);

            // Call once.
            OnTimer(null, null);
        }

        protected override void OnStop()
        {
            eventLog.Close();
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            // Check for all overdue invoices.
            eventLog.WriteEntry("Checking for overdue invoices...", EventLogEntryType.Information, eventId++);

            // Connect to database.
            string connectionString = @"Persist Security Info=False;server=10.1.1.17;User ID=ar;Password=356tCS716^;";
            using (SqlConnection cnn = new SqlConnection(connectionString))
            {
                // First, figure out how many invoices need to be processed.
                string query = $"SELECT COUNT(*) FROM [Vision].[dbo].[inMaster] WHERE Posted = 'Y' AND (DATEDIFF(day, TransDate, GETDATE()) = {reminder_time} OR DATEDIFF(day, TransDate, GETDATE()) = {overdue_time});";
                SqlCommand cmd = new SqlCommand(query, cnn);

                try
                {
                    cnn.Open();
                }
                catch (Exception ex)
                {
                    eventLog.WriteEntry("Couldn't connect to database: " + ex.Message, EventLogEntryType.Error, eventId++);

                    // Send error email to admin!

                }

                SqlDataReader reader = cmd.ExecuteReader();
                Int32 aged = 0;

                try
                {
                    while (reader.Read())
                    {
                        aged = reader.GetInt32(0);   
                    }
                }
                catch (Exception ex)
                {
                    eventLog.WriteEntry("Couldn't read from database: " + ex.Message, EventLogEntryType.Error, eventId++);
                    
                    // Send error email to admin!
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }

                // Only process daily if there are accounts to be processed. Otherwise, this skips to the end.
                if (aged > 0)
                {
                    eventLog.WriteEntry("Found " + aged + " aged accounts", EventLogEntryType.Information, eventId++);
                    // Get relevant info (job number, invoice number, due date, posted, late).
                    query = $"SELECT WBS1, Invoice, DATEDIFF(day, TransDate, GETDATE()) as DueTime FROM [Vision].[dbo].[inMaster] WHERE Posted = 'Y' AND (DATEDIFF(day, TransDate, GETDATE()) = {reminder_time} OR DATEDIFF(day, TransDate, GETDATE()) = {overdue_time});";
                    cmd = new SqlCommand(query, cnn);


                    // Stores Job object with job number as key.
                    Dictionary<string, Job> jobDictionary = new Dictionary<string, Job>();

                    reader = cmd.ExecuteReader();

                    try
                    {
                        while (reader.Read())
                        {
                            // Extract data from query.
                            if (!jobDictionary.ContainsKey(reader["WBS1"].ToString()))
                                jobDictionary[reader["WBS1"].ToString()] = new Job(id: reader["WBS1"].ToString(), late: Convert.ToInt32(reader["DueTime"]) == overdue_time);

                            jobDictionary[reader["WBS1"].ToString()].invoices.Add(new Invoice(reader["Invoice"].ToString(), reader["WBS1"].ToString()));
                        }
                    } catch (Exception ex)
                    {
                        eventLog.WriteEntry("Couldn't read job information from database: " + ex.Message, EventLogEntryType.Error, eventId++);
                    } finally
                    {
                        reader.Close();
                    }

                    Dictionary<string, Client> clients = new Dictionary<string, Client>();
                    List<string> foundClients = new List<string>();

                    // First, check to see if the client has paid already. We don't want to send a reminder if they have.
                    foreach (var pair in jobDictionary)
                    {
                        query = @"SELECT inMaster.WBS1, SUM(CASE WHEN LedgerAR.TransType = 'IN' THEN ABS(LedgerAR.Amount) END) as Owed, SUM(CASE WHEN LedgerAR.TransType = 'CR' THEN LedgerAR.Amount END) as Paid
                                    FROM [Vision].[dbo].[inMaster]
                                    LEFT JOIN LedgerAR ON inMaster.WBS1 = LedgerAR.WBS1 
                                    WHERE inMaster.WBS1 = @job
                                    GROUP BY inMaster.WBS1";

                        cmd = new SqlCommand(query, cnn);
                        cmd.Parameters.AddWithValue("@job", pair.Key);

                        reader = cmd.ExecuteReader();

                        bool found = false;

                        try
                        {
                            while (reader.Read())
                            {
                                // 2nd value will be null if no payment has been made.
                                if (reader.GetValue(2) != DBNull.Value)
                                {
                                    // If paid and due are equal, invoice has been paid and we can set found flag.
                                    if ((reader.GetValue(1) as int? ?? 0) + (reader.GetValue(2) as int? ?? 0) == 0)
                                    {
                                        // already paid.
                                        found = true;
                                        eventLog.WriteEntry("Already paid, skipping...", EventLogEntryType.Information, eventId++);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            eventLog.WriteEntry("Couldn't read invoice information from database: " + ex.Message + ". Emailing accounting/ccs", EventLogEntryType.Error, eventId++);
                            mailer.SendMail("Unable to read invoice file from database", "accounting@augustmack.com", "AutomateAR was unable to read the invoice for job " + pair.Key + ".", null, ccAddress);
                        }
                        finally
                        {
                            reader.Close();
                        }

                        if (found)
                        {
                            foundClients.Add(pair.Key);
                            continue;
                        }
                    }

                    // Remove client with no bill due.
                    foreach (var clientKey in foundClients)
                    {
                        jobDictionary.Remove(clientKey);
                    }

                    foreach (var pair in jobDictionary)
                    {
                        foreach (var inv in pair.Value.invoices)
                        {
                            string invoice = inv.id;

                            // Get invoice.
                            query = "SELECT FinalInvoiceFileID, EditedFileType FROM [Vision].[dbo].[billInvMaster] WHERE Invoice = @Invoice;";
                            cmd = new SqlCommand(query, cnn);

                            cmd.Parameters.AddWithValue("@Invoice", invoice);


                            reader = cmd.ExecuteReader();

                            try
                            {
                                while (reader.Read())
                                {
                                    inv.file_id = reader["FinalInvoiceFileID"].ToString();
                                    inv.extension = reader["EditedFileType"].ToString();
                                }
                            }
                            catch (Exception ex)
                            {
                                eventLog.WriteEntry("Couldn't read invoice file information from database: " + ex.Message + " Emailing accounting/ccs.", EventLogEntryType.Error, eventId++);
                                mailer.SendMail("Unable to read invoice file from database", "accounting@augustmack.com", "AutomateAR was unable to read the invoice for job " + pair.Key + ".", null, ccAddress);
                                pair.Value.invoices.Clear();
                            }
                            finally
                            {
                                reader.Close();
                            }


                            // Get invoice file data.
                            query = "SELECT FileData FROM [VisionFILES].[dbo].[FW_FILES] WHERE FileID = @FileID";
                            cmd = new SqlCommand(query, cnn);

                            cmd.Parameters.AddWithValue("@FileID", inv.file_id);


                            reader = cmd.ExecuteReader();


                            // Extract pdf data from database.
                            try
                            {
                                while (reader.Read())
                                {
                                    inv.invoice_data = (byte[]) reader["FileData"];
                                }
                            }
                            catch (Exception ex)
                            {
                                eventLog.WriteEntry("Couldn't read invoice data from database: " + ex.Message + " Emailing accounting/ccs.", EventLogEntryType.Error, eventId++);
                                mailer.SendMail("Unable to read invoice file from database", "accounting@augustmack.com", "AutomateAR was unable to read the invoice for job " + pair.Key + ".", null, ccAddress);
                                pair.Value.invoices.Clear();
                            }
                            finally
                            {
                                reader.Close();
                            }
                        }
                    
                        // Get contact and send an email.
                        query = "SELECT BillingContactID FROM [Vision].[dbo].[PR] WHERE WBS1 = @WBS1;";
                        cmd = new SqlCommand(query, cnn);

                        cmd.Parameters.AddWithValue("@WBS1", pair.Key);

                        reader = cmd.ExecuteReader();

                        try
                        {
                            while (reader.Read())
                            {
                                if (!clients.ContainsKey(reader["BillingContactID"].ToString()))
                                {
                                    clients[reader["BillingContactID"].ToString()] = new Client(reader["BillingContactID"].ToString());

                                    pair.Value.contactID = reader["BillingContactID"].ToString();
                                    clients[reader["BillingContactID"].ToString()].jobs.Add(pair.Value);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            eventLog.WriteEntry("Couldn't read contact information from database: " + ex.Message, EventLogEntryType.Error, eventId++);
                        }
                        finally
                        {
                            reader.Close();
                        }


                        query = "SELECT EMail FROM [Vision].[dbo].[Contacts] WHERE ContactID = @ContactID;";
                        cmd = new SqlCommand(query, cnn);

                        cmd.Parameters.AddWithValue("@ContactID", pair.Value.contactID);

                        reader = cmd.ExecuteReader();

                        try
                        {
                            while (reader.Read())
                            {
                                clients[pair.Value.contactID].email = reader["EMail"].ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            eventLog.WriteEntry("Couldn't read contact from database: " + ex.Message, EventLogEntryType.Error, eventId++);
                        }
                        finally
                        {
                            reader.Close();
                        }
                    }

                    eventLog.WriteEntry(jobDictionary.Count.ToString(), EventLogEntryType.Information, eventId++);

                    // Actually send email.

                    foreach (var pair in clients)
                    {
                        if (!dryRun && allowedAddress.Count() != 0 && allowedAddress.Contains(pair.Value.email))
                            mailer.SendMail(pair.Value.jobs[0].late ? "Payment Past Due" : "Friendly Reminder for Open Invoice(s)", pair.Value.email, mailer.GenerateEmail(pair.Value.jobs[0].late ? EmailType.Due : EmailType.Reminder, pair.Value.email), pair.Value.jobs[0].invoices[0].invoice_data, ccAddress);
                        else
                            mailer.SendMail(pair.Value.jobs[0].late ? "Payment Past Due" : "Friendly Reminder for Open Invoice(s)", "accounting@augustmack.com", mailer.GenerateEmail(pair.Value.jobs[0].late ? EmailType.Due : EmailType.Reminder, pair.Value.email), pair.Value.jobs[0].invoices[0].invoice_data, ccAddress);
                    }
                } else
                {
                    
                    eventLog.WriteEntry("No aged invoices -- skipping.", EventLogEntryType.Information, eventId++);
                }
            }
        }
    }
}
