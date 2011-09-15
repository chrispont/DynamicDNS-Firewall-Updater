// <copyright file="Service1.cs" company="Chris Pont">
// Copyright 2011, Chris Pont
// [Copyright]
// </copyright>

namespace FirewallService
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Management.Automation;
    using System.Management.Automation.Runspaces;
    using System.Net;
    using System.ServiceProcess;
    using System.Text;
    using System.Threading;

    /// <summary>
    /// The Firewall Service
    /// </summary>
    public partial class Service1 : ServiceBase
    {
        /// <summary>
        /// The last IP address found
        /// </summary>
        private string lastIPAddress = string.Empty;

        /// <summary>
        /// Configuration Settings Object
        /// </summary>
        private ConfigurationSettings config;

        /// <summary>
        /// The timer which fires to check the IP Address
        /// </summary>
        private System.Timers.Timer timer1;

        /// <summary>
        /// A mutex object
        /// </summary>
        private Mutex mutex;

        /// <summary>
        /// Initializes a new instance of the <see cref="Service1"/> class.
        /// </summary>
        public Service1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// When implemented in a derived class, executes when a Start command is sent to the service by the Service Control Manager (SCM) or when the operating system starts (for a service that starts automatically). Specifies actions to take when the service starts.
        /// </summary>
        /// <param name="args">Data passed by the start command.</param>
        protected override void OnStart(string[] args)
        {
            System.Configuration.AppSettingsReader appReader = new System.Configuration.AppSettingsReader();
            this.timer1 = new System.Timers.Timer(Convert.ToDouble(appReader.GetValue("Interval", typeof(string))));

            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = System.AppDomain.CurrentDomain.BaseDirectory;
            /* Watch for changes in LastAccess and LastWrite times, and
               the renaming of files or directories. */
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
               | NotifyFilters.FileName | NotifyFilters.DirectoryName;

            // Only watch text files.
            watcher.Filter = "*.config";

            // Add event handlers.
            watcher.Changed += new FileSystemEventHandler(OnChanged);

            // Begin watching.
            watcher.EnableRaisingEvents = true;
            this.mutex = new Mutex(false);
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.Timer1_Elapsed);
            this.timer1.Start();

            string ipToAdd = appReader.GetValue("Host", typeof(string)).ToString();
            EventLog.WriteEntry("Firewall Updater monitoring host '" + ipToAdd + "'");
        }

        /// <summary>
        /// When implemented in a derived class, executes when a Stop command is sent to the service by the Service Control Manager (SCM). Specifies actions to take when a service stops running.
        /// </summary>
        protected override void OnStop()
        {
            EventLog.WriteEntry("Firewall Updater exiting.");
        }

        /// <summary>
        /// Called when [changed].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="e">The <see cref="System.IO.FileSystemEventArgs"/> instance containing the event data.</param>
        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            ConfigurationManager.RefreshSection("appSettings");
        }

        /// <summary>
        /// Handles the Elapsed event of the Timer1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Timers.ElapsedEventArgs"/> instance containing the event data.</param>
        private void Timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            this.mutex.WaitOne();
            try
            {
                // EventLog.WriteEntry("Tick");
                System.Configuration.AppSettingsReader appReader = new System.Configuration.AppSettingsReader();
                string ipToAdd = appReader.GetValue("Host", typeof(string)).ToString();
                IPAddress[] addresslist = Dns.GetHostAddresses(ipToAdd);

                if (addresslist.Length > 0)
                {
                    if (this.lastIPAddress != addresslist[0].ToString())
                    {
                        this.RunScript("netsh advfirewall firewall delete rule name=\"" + appReader.GetValue("RuleName", typeof(string)).ToString() + "\" protocol=TCP");
                        string scripttext = "netsh advfirewall firewall add rule name=\"" + appReader.GetValue("RuleName", typeof(string)).ToString() + "\" dir=in action=allow protocol=TCP remoteip=" + addresslist[0];
                        this.RunScript(scripttext);
                        EventLog.WriteEntry("Changed firewall rule for RDP. New IP Address: " + addresslist[0]);
                        this.lastIPAddress = addresslist[0].ToString();
                    }
                }
            }
            finally
            {
                this.mutex.ReleaseMutex();
            }
        }

        /// <summary>
        /// Runs the script.
        /// </summary>
        /// <param name="scriptText">The script text.</param>
        /// <returns>The output of the script run</returns>
        private string RunScript(string scriptText)
        {
            // create Powershell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();

            // open it
            runspace.Open();

            // create a pipeline and feed it the script text
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript(scriptText);
            pipeline.Commands.Add("Out-String");

            // execute the script
            Collection<PSObject> results = pipeline.Invoke();

            // close the runspace
            runspace.Close();

            // convert the script result into a single string
            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject obj in results)
            {
                stringBuilder.AppendLine(obj.ToString());
            }

            return stringBuilder.ToString();
        }
    }
}
