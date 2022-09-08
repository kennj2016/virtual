// =====================================================================================================
// = Program: SusplanReportMonitor  Version: 2.6.2.1
// =====================================================================================================
// = Created: 07/28/2006                                                           
// = Author: John Donehoo                                                          
// = Description: This is a Windows Service that monitors the completion of  
// =              Adlib eXpress PDF generation. 
// = 
// = ----------------------------------------------------------------------------  
// = Change Log                                                                    
// =                                                                               
// =    Date      Version  In  Description                                             
// = 07/28/2006  02.03.03  jd  New Program - the old VB5 ReportMonitor program 
// =                           rewritten as a VB .Net windows service.
// = 06/05/2007  02.04.00  jd  Add support for new directory structure
// = 10/14/2008  2.5.1.5   jd  When a document is added set the update date in addition to the create date.
// = 01/23/2009  2.5.1.7   jd  Fix: A nonstandard documents directory was not being properly handled.
// = 11/06/2009  2.6.1.4   jd  Handle single quotes in target (physical) name on document update.
// =====================================================================================================

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.ServiceProcess;
using Microsoft.Win32;
using ServiceDebuggerHelper;

namespace ReportProcess
{
    public partial class SusplanReportMonitor : DebuggableService
    {
        public bool verbose;
        private List<SPInstanceRM> collInstances; // collection of SPInstanceRM objects

        // the variables below are loaded from the registry
        private long sleepSeconds;   // sleep seconds (report queue sweep interval)
        private long longSleepSeconds;  // long sleep seconds (report queue purge processing interval)
        public int completedPurgeDays; // age in days for "completed" queue entries to be purged
        public int cancelledPurgeDays; // age in days for "cancelled" queue entries to be purged
        public int errorPurgeDays;     // age in days for "error" queue entries to be purged
        public int longPendingPurgeDays; // age in days for "pending", "crystal processing", and "in progress" queue entries to be purged

        // defaults for the above variables
        private const long DEFAULTSLEEPSECONDS = 15L;
        private const long DEFAULTLONGSLEEPSECONDS = 600L;
        private const int DEFAULTCOMPLETEDPURGEDAYS = 1;
        private const int DEFAULTCANCELLEDPURGEDAYS = 1;
        private const int DEFAULTERRORPURGEDAYS = 1;
        private const int DEFAULTLONGPENDINGPURGEDAYS = 1;

        // registry key for operational parameters
        private const string OPERATIONALPARMSRK = @"Software\Virtual Corporation\Sustainable_Planner\Services\ReportMonitor";

        // registry key for SP instances
        private const string INSTANCEPARMSRK = @"Software\Virtual Corporation\Sustainable_Planner\Instances\";
        private const string myVersion = "4.0.0";
        private const string myVersionDate = "07/01/2022";


        /* TODO ERROR: Skipped RegionDirectiveTrivia */
        public SusplanReportMonitor() : base()
        {

            // This call is required by the Component Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call

        }

        // UserService overrides dispose to clean up the component list.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components is object)
                {
                    components.Dispose();
                }
            }

            base.Dispose(disposing);
        }

        // The main entry point for the process
        [MTAThread()]
        public static void Main(string[] args)
        {
            ServiceBase[] ServicesToRun;

            // More than one NT Service may run within the same process. To add
            // another service to this process, change the following line to
            // create a second service object. For example,
            // 
            // ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
            // 
            if (args.Length > 0 && args[0].ToLower().Equals("/debug"))
            {
                System.Windows.Forms.Application.Run(new ServiceRunner(new SusplanReportMonitor()));
            }
            else
            {
                ServicesToRun = new ServiceBase[] { new SusplanReportMonitor() };
                ServicesToRun[0].CanShutdown = true;
                ServiceBase.Run(ServicesToRun);
            }
        }

        // Required by the Component Designer
        private System.ComponentModel.IContainer components = null;

        // NOTE: The following procedure is required by the Component Designer
        // It can be modified using the Component Designer.  
        // Do not modify it using the code editor.
        internal EventLog evtEventLog1;
        private System.Timers.Timer _tmrTimer1;

        internal System.Timers.Timer tmrTimer1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _tmrTimer1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_tmrTimer1 != null)
                {
                    _tmrTimer1.Elapsed -= tmrTimer1_Elapsed;
                }

                _tmrTimer1 = value;
                if (_tmrTimer1 != null)
                {
                    _tmrTimer1.Elapsed += tmrTimer1_Elapsed;
                }
            }
        }

        private System.Timers.Timer _tmrTimerLong;

        internal System.Timers.Timer tmrTimerLong
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _tmrTimerLong;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_tmrTimerLong != null)
                {
                    _tmrTimerLong.Elapsed -= tmrTimerLong_Elapsed;
                }

                _tmrTimerLong = value;
                if (_tmrTimerLong != null)
                {
                    _tmrTimerLong.Elapsed += tmrTimerLong_Elapsed;
                }
            }
        }

        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            evtEventLog1 = new EventLog();
            _tmrTimer1 = new System.Timers.Timer();
            _tmrTimer1.Elapsed += new System.Timers.ElapsedEventHandler(tmrTimer1_Elapsed);
            _tmrTimerLong = new System.Timers.Timer();
            _tmrTimerLong.Elapsed += new System.Timers.ElapsedEventHandler(tmrTimerLong_Elapsed);
            ((System.ComponentModel.ISupportInitialize)evtEventLog1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)_tmrTimer1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)_tmrTimerLong).BeginInit();
            // 
            // evtEventLog1
            // 
            evtEventLog1.Log = "Application";
            evtEventLog1.Source = "SusplanReportMonitor";
            // 
            // tmrTimer1
            // 
            // 
            // tmrTimerLong
            // 
            // 
            // SusplanReportMonitor
            // 
            CanPauseAndContinue = true;
            CanShutdown = true;
            ServiceName = "SusplanReportMonitor";
            ((System.ComponentModel.ISupportInitialize)evtEventLog1).EndInit();
            ((System.ComponentModel.ISupportInitialize)_tmrTimer1).EndInit();
            ((System.ComponentModel.ISupportInitialize)_tmrTimerLong).EndInit();
        }

        /* TODO ERROR: Skipped EndRegionDirectiveTrivia */
        protected override void OnStart(string[] args)
        {
            // Add code here to start your service. This method should set things
            // in motion so your service can do its work.

            // --------------------------------------------------
            // Read command line parameters and switches
            // --------------------------------------------------

            string argparmCode = "-v";
            verbose = GetParmExists(ref argparmCode, args);
            LogMsg("SusplanReportMonitor Starting", false);
            LogMsg("Version: " + myVersion + " " + myVersionDate, false);

            if (verbose)
            {
                LogMsg("Verbose Messages in effect (see event log)", false);
            }

            // initialize variables
            collInstances = new List<SPInstanceRM>();  // initialize the collection of SPInstances

            // read operational parameters from the registry
            ReadOperationalParms();

            // build the collection of SPInstances
            ReadDBParms();      // load the database parms for each SP Instance
            if (collInstances.Count == 0)
            {
                string argerrMsg = "No valid SP instances found.";
                DoTerminate(ref argerrMsg);
            }

            LogMsg("About to do first (short) timer waits.", true);
            tmrTimer1.Interval = 500d;    // set short timer interval in milliseconds
            tmrTimer1.Enabled = true;
            tmrTimer1.AutoReset = true;
        }

        protected override void OnStop()
        {
            // Add code here to perform any tear-down necessary to stop your service.
            tmrTimer1.Enabled = false;
            tmrTimerLong.Enabled = false;
            LogMsg("SusplanReportMonitor Stopping", false);

            LogMsg("SusplanReportMonitor Stopped", false);
        }

        protected override void OnShutdown()
        {
            tmrTimer1.Enabled = false;
            tmrTimerLong.Enabled = false;
            LogMsg("System Shutdown Received - SusplanReportMonitor Stopping", false);

            LogMsg("SusplanReportMonitor Terminated", false);
        }

        private void tmrTimer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            tmrTimer1.Enabled = false;
            LogMsg("Sleep Timer Popped - About to perform report sweep", true);
            ProcessPendingAllInstances();          // Perform initial processing of instances
            LogMsg("Going back to sleep", true);
            tmrTimer1.Interval = sleepSeconds * 1000L;  // set timer interval in milliseconds
            tmrTimer1.Enabled = true;
        }

        private void tmrTimerLong_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            tmrTimerLong.Enabled = false;
            LogMsg("Long Sleep Timer Popped", true);
            ProcessPurgeAllInstances();  // Perform purge processing for all instances 
            LogMsg("Going back to long sleep", true);
            tmrTimerLong.Interval = longSleepSeconds * 1000L;   // set timer interval in milliseconds
            tmrTimerLong.Enabled = true;
        }

        protected override void OnPause()
        {
            tmrTimer1.Enabled = false;
            tmrTimerLong.Enabled = false;
            LogMsg("SusplanReportMonitor Paused", false);
        }

        protected override void OnContinue()
        {
            LogMsg("SusplanReportMonitor Continuing", false);
            ProcessPendingAllInstances();    // Perform report processing for all instances
            ProcessPurgeAllInstances();      // Perform purge processing for all instances 
            tmrTimer1.Interval = sleepSeconds * 1000L;    // set timer interval in milliseconds
            tmrTimer1.Enabled = true;
            tmrTimerLong.Interval = longSleepSeconds * 1000L; // set timer interval in milliseconds
            tmrTimerLong.Enabled = true;
        }

        // ------------------------------------------------------------------------------
        // Process any pending reports for all instances
        // ------------------------------------------------------------------------------
        private void ProcessPendingAllInstances()
        {
            LogMsg("Start processing reports for all instances", true);
            foreach (SPInstanceRM mySPInstance in collInstances)
            {
                try
                {
                    mySPInstance.SweepReports();
                }
                catch (Exception ex)
                {
                    evtEventLog1.WriteEntry("Report Queue Exception Error : " + ex.ToString(), EventLogEntryType.Error);
                } // Error message in Event log
            }

            LogMsg("Finished processing reports for all instances", true);
        }


        // ------------------------------------------------------------------------------
        // Purge the report queue database table for old and/or obsolete entries
        // ------------------------------------------------------------------------------
        private void ProcessPurgeAllInstances()
        {
            LogMsg("Start purge processing for all instances", true);
            foreach (SPInstanceRM mySPInstance in collInstances)
            {
                LogMsg("Start purge processing for instance:" + mySPInstance.InstanceName, true);
                try
                {
                    mySPInstance.CheckCompleted();  // delete entries completed for x days
                    mySPInstance.CheckCancelled();  // delete entries cancelled for x days
                    mySPInstance.CheckErrors();     // delete entries in error status for x days
                    mySPInstance.CheckLongPending();    // delete entries stuck for x days
                    mySPInstance.FlushLogFile();        // flush log file entries to disk
                }
                catch (Exception ex)
                {
                    evtEventLog1.WriteEntry("Queue Purge Processing Exception Error : " + ex.ToString(), EventLogEntryType.Error);
                }   // Error message in Event log
            }

            LogMsg("Finished processing completed reports for all instances", true);
        }

        // ------------------------------------------------------------------------------
        // Read operational parameters from the registry
        // ------------------------------------------------------------------------------
        private void ReadOperationalParms()
        {
            RegistryKey rk;
            LogMsg("Reading Operational Parameters", false);
            rk = Registry.LocalMachine.OpenSubKey(OPERATIONALPARMSRK);

            // sleep seconds
            if (rk is null)                       // missing?
            {
                sleepSeconds = DEFAULTSLEEPSECONDS;  // yep - use default
                LogMsg("Sleep Seconds Not Found in Registry - Using Default Value", false);
            }
            else
            {
                sleepSeconds = Convert.ToInt32(rk.GetValue("SleepSeconds", DEFAULTSLEEPSECONDS.ToString()));
            }

            LogMsg("Sleep Seconds = " + sleepSeconds.ToString(), false);

            // long sleep seconds 
            if (rk is null)                           // missing?
            {
                longSleepSeconds = DEFAULTLONGSLEEPSECONDS;    // yep - use default
                LogMsg("Long Sleep Seconds Not Found in Registry - Using Default Value", false);
            }
            else
            {
                longSleepSeconds = Convert.ToInt32(rk.GetValue("LongSleepSeconds", DEFAULTLONGSLEEPSECONDS.ToString()));
            }

            LogMsg("Long Sleep Seconds = " + longSleepSeconds.ToString(), false);

            // -----------------------------------------------------------
            // "completed" purge days
            // -----------------------------------------------------------
            if (rk is null)                               // missing?
            {
                completedPurgeDays = DEFAULTCOMPLETEDPURGEDAYS;  // yep - use default
                LogMsg("Completed Entry Purge Days Not Found in Registry - Using Default Value", false);
            }
            else
            {
                completedPurgeDays = Convert.ToInt32(rk.GetValue("CompletedPurgeDays", DEFAULTCOMPLETEDPURGEDAYS.ToString()));
            }

            LogMsg("Completed Entry Purge Days = " + completedPurgeDays.ToString(), false);

            // -----------------------------------------------------------
            // "cancelled" purge days
            // -----------------------------------------------------------
            if (rk is null)                               // missing?
            {
                cancelledPurgeDays = DEFAULTCANCELLEDPURGEDAYS;  // yep - use default
                LogMsg("Cancelled Entry Purge Days Not Found in Registry - Using Default Value", false);
            }
            else
            {
                cancelledPurgeDays = Convert.ToInt32(rk.GetValue("CancelledPurgeDays", DEFAULTCANCELLEDPURGEDAYS.ToString()));
            }

            LogMsg("Cancelled Entry Purge Days = " + cancelledPurgeDays.ToString(), false);

            // -----------------------------------------------------------
            // "error" purge days
            // -----------------------------------------------------------
            if (rk is null)                       // missing?
            {
                errorPurgeDays = DEFAULTERRORPURGEDAYS;  // yep - use default
                LogMsg("Error Entry Purge Days Not Found in Registry - Using Default Value", false);
            }
            else
            {
                errorPurgeDays = Convert.ToInt32(rk.GetValue("ErrorPurgeDays", DEFAULTERRORPURGEDAYS.ToString()));
            }

            LogMsg("Error Entry Purge Days = " + errorPurgeDays.ToString(), false);

            // -----------------------------------------------------------
            // "long pending" purge days
            // -----------------------------------------------------------
            if (rk is null)                               // missing?
            {
                longPendingPurgeDays = DEFAULTLONGPENDINGPURGEDAYS;  // yep - use default
                LogMsg("Long Pending Entry Purge Days Not Found in Registry - Using Default Value", false);
            }
            else
            {
                longPendingPurgeDays = Convert.ToInt32(rk.GetValue("LongPendingPurgeDays", DEFAULTLONGPENDINGPURGEDAYS.ToString()));
            }

            LogMsg("Long Pending Entry Purge Days = " + longPendingPurgeDays.ToString(), false);
            if (rk is object)
            {
                rk.Close();
            }

            LogMsg("Finished Reading Operational Parameters", false);
        }

        // ------------------------------------------------------------------------------
        // Retrieve the location of the configuration file for each Sustainable Planner
        // instance from the Windows Registry.
        // This file will be used to initialize an SPInstanceRM object.
        // ------------------------------------------------------------------------------

        private void ReadDBParms()
        {
            var susplanReportMonitor = this;
            string instanceName = ConfigurationManager.AppSettings["InstanceName"].ToString();
            SPInstanceRM mySPInstance = new SPInstanceRM(instanceName, ref susplanReportMonitor);
            if (!mySPInstance.isValid)     // error trying to create it?
            {
                LogMsg("Instance Test will not be processed.", false, EventLogEntryType.Error);
                mySPInstance = default;
            }
            else
            {
                collInstances.Add(mySPInstance);
            }
        }

        // --------------------------------------------------
        // Routines for command line argument processing
        // --------------------------------------------------

        // return true if a given parm is on the command line
        public bool GetParmExists(ref string parmCode, string[] args)
        {
            foreach (var myArg in args)
            {
                if ((myArg ?? "") == (parmCode ?? ""))
                {
                    return true;
                }
            }

            return false;
        }

        // --------------------------------------------------
        // Abnormal Termination Routine
        // --------------------------------------------------
        public void DoTerminate(ref string errMsg)
        {
            LogMsg(errMsg, false);
            LogMsg("******Processing Terminated******", false, EventLogEntryType.Error);
            Environment.Exit(0);
        }

        // --------------------------------------------------
        // Write a message to the event log
        // --------------------------------------------------
        // log an information type message
        public void LogMsg(string msg, bool debugMsg)
        {
            DoLogMsg(msg, debugMsg, EventLogEntryType.Information);  // information message
        }

        // log a message using the passed log entry type
        public void LogMsg(string msg, bool debugMsg, EventLogEntryType logEntryType)

        {
            DoLogMsg(msg, debugMsg, logEntryType);
        }

        public void DoLogMsg(string msg, bool debugMsg, EventLogEntryType logEntryType)

        {
            string wkMsg;
            if (debugMsg)
            {
                wkMsg = ">>> " + msg;
            }
            else
            {
                wkMsg = msg;
            }

            if (!debugMsg | verbose)
            {
                try
                {
                    evtEventLog1.WriteEntry(wkMsg, logEntryType);
                }
                catch
                {
                    // no action - we don't want to hang up the service if there
                    // is an event log error
                }
            }
        }
    }
}