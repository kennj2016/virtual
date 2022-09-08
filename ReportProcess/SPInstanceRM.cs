using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using CrystalDecisions.CrystalReports.Engine;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using ReportProcess.DBContexts;
using System.Data.Entity;
using System.Configuration;
using CrystalDecisions.Shared;

namespace ReportProcess
{
    public partial class SPInstanceRM
    {
        #region Properties
        private readonly UNIT_TESTContext _context = new UNIT_TESTContext();
        private readonly string instName;
        private readonly string documentPath;
        private readonly SusplanReportMonitor myCaller;
        private bool logFileInitialized;
        private string logFileName;
        private string logFilePath;
        private int logFileDay;
        private bool logFileValid;
        private static EventLog objEventLog;
        private FileStream objFSlog;
        private StreamWriter objSWlog;
        private List<string> collLogFileQueue;
        public bool isValid;

        // the variables below are for the report queue entry being processed
        private string svTargetDirectory, svTargetName;
        private int svRQId, svTOCId, svNodeId;
        private int rc;
        private int qSNextSectionId;
        private int qSNFirstChildId;
        private int qSFirstTOCDOCId;
        private string qSSectionName;
        private int fwkSecOption;
        private string svInputDocInd;
        private int qSectionId;
        private int rpct = 0;
        private int qTocDocId;
        // Each entry Consists of:
        // 1. filename
        // 2. label
        // 3. level
        // 4. file path
        // 5. section name
        private string[,] rptPiecesTable = new string[201, 6];  // Report Pieces Table
        private int qNextTocDocId;
        private string qTocDocLabel;
        private int ntdx;
        private int qNodeId;
        private string qDocLabel;
        private readonly int[] nodeTbl = new int[11]; // table of nodes from the current node to the top (entry 0 is the active entry count)
        private string svErrorMsg;
        private string svTempDir;
        private readonly string svRQId9;
        private string svDocDirId;
        private string _nodeDocDirTbl;
        private string svCrystalInd;
        private string svErrorInd;
        private int svFirstSectionID;
        private string svDateTime;
        private string qNDescription;
        private int qNParentId;
        private int qPreviousNodeId;
        private int qNextNodeId;
        private int qNFirstChildId;
        private int qLevelNumber;
        private string qDocDirId;
        private string svNodeName;
        private string svRptFileName;
        private string svTOCReportName;
        private string tempDir;
        private string svCtlFileName;
        private int rpdx;
        private string wkFileName;
        private string wkFileLabel;
        private string wkFileDir;
        private bool gotError;
        private string svSurveyId;
        private string svFooterYN;
        private string svFooterFont;
        private int svTOCPage;
        private string svTocLayout;
        private string svTocLeftMargin;
        private string svTocRightMargin;
        private string svTocFont;
        private string svTocFontSize;
        private int svHeaderStartOnDoc;
        private string svHeaderFont;
        private string svHeaderFontSize;
        private string svFooterFontSize;
        private string svHeaderText;
        private int svFooterStartOnDoc;
        private string svFooterL;
        private string svFooterC;
        private string svFooterR;
        private string svStdPageNumbers;
        private int svTOCFirstSectionId;
        private char svCoverYN;
        private int svCvrPageID;
        private string svCvrPage;
        private char svTocYN;
        private string svFootLMargin;
        private string svFootRMargin;
        private string svFootBMargin;
        private string svFooterOptL;
        private string svFooterOptR;
        private string svFooterOptC;
        private int svTOCNodeId;
        private char svWaterYN;
        // Private evtEventLog1 As System.Diagnostics.EventLog

        // ------------------------------------------------------------------------------
        // Properties
        // ------------------------------------------------------------------------------

        // return the instance name
        public string InstanceName
        {
            get
            {
                return instName;
            }
        }

        #endregion

        // ------------------------------------------------------------------------------
        // Initialize a new SPInstanceRM object.
        // The location of the SPParms.spp file is passed to this subroutine.
        // Data in the configuration file is used to validate the product license key.
        // Database access parameters are also retrieved and the database connection
        // is opened.
        // The location of the documents directory(s)is loaded from the system parameters 
        // database table or derived from the location of the configuration file.
        // ------------------------------------------------------------------------------
        public SPInstanceRM(string instanceName, ref SusplanReportMonitor caller)

        {
            myCaller = caller;
            instName = instanceName;
            isValid = false;        // not a valid instance
            logFileInitialized = false; // log file not initialized
            logFileValid = false;       // log file is not available
            collLogFileQueue = new List<string>();  // collection holds messages till file is open
            objEventLog = myCaller.evtEventLog1;    // store event log object

            // --------------------------------------------------
            // Get documents directory information for the 
            // database
            // --------------------------------------------------
            LogMsg("Accessing System Parameters Database Record", true);
            SUSPLAN_PARMS param = _context.SUSPLAN_PARMS.Where(p => p.ID == 0).FirstOrDefault();
            string fwkDocStdLocationYN;
            if (param != null)
            {
                documentPath = param.DOC_DIRECTORY;
                fwkDocStdLocationYN = param.DOC_STD_LOCATION_YN;
                if (string.IsNullOrEmpty(fwkDocStdLocationYN))
                {
                    fwkDocStdLocationYN = "y";
                }
            }
            else
            {
                LogMsg("System Parameters database record not found.", false, EventLogEntryType.Error);
                return;
            }

            if (string.IsNullOrEmpty(documentPath))
            {
                LogMsg("No path to documents directory", false, EventLogEntryType.Error);
                return;
            }

            LogMsg("Path to Documents Directory: " + documentPath, true);
            isValid = true;
            OpenLogFile();  // open the disk log file and write any queued messages
            LogMsg("Initialization Complete for Instance: " + instName, false);
        }

        private string Get_nodeDocDirTbl(int ntdx)
        {
            return _nodeDocDirTbl;
        }

        private void Set_nodeDocDirTbl(int ntdx, string value)
        {
            _nodeDocDirTbl = value;
        }

        // -----------------------------------------------------------------------------
        // Process any pending reports
        // Read all "pending" queue entries
        // look for the "target" file name in the "target directory"
        // If it exists the pdf processing is done.
        // Call the "ProcessPending" routine to set to the cleanup processing
        // and mark the entry "Complete".
        // If the entry is not complete, check the pdf_errors directory to see of
        // adlib has moved the job ticket there.  If it has, tag the queue entry with
        // status = "Error Encountered", and rename the job ticket
        // -----------------------------------------------------------------------------

        public void SweepReports()
        {
            // Dim wkXMLFileName As String
            LogMsg("Starting Report Queue Sweep", true);

            FlushLogFile(); // flush log file entries to disk
            List<SUSPLAN_REPORT_Q> listReportQ = _context.SUSPLAN_REPORT_Q.Where(p => p.STATUS == "P").OrderBy(p => p.RQ_ID).ToList();
            foreach (var item in listReportQ)
            {
                svRQId = item.RQ_ID;
                svTargetDirectory = item.TARGET_DIRECTORY;
                svTargetName = item.TARGET_NAME;
                svTOCId = item.TOC_ID ?? 0;
                svNodeId = item.NODE_ID ?? 0;
                svFirstSectionID = item.FIRST_SEC_ID ?? 0;
                svSurveyId = item.SURVEY_ID.HasValue ? item.SURVEY_ID.Value.ToString() : "";
                svTempDir = documentPath + @"\pdf_output\";
                rptPiecesTable = new string[201, 6];
                Array.Clear(rptPiecesTable, 0, rptPiecesTable.Length);
                rpct = 0; // initialize report pieces count
                rpdx = 0; // initialize index
                ProcessSection(svFirstSectionID, 1);
                if (svInputDocInd == "N")   // was there at least one input document?	
                {
                    svErrorInd = "M"; // no - set error indicator
                }

                if (svInputDocInd == "Y")   // any input documents?
                {
                    createRptCtlFileXML(); // write the report control file (XML Style)
                }
            }

            LogMsg("Ending Report Queue Sweep", true);
        }

        // ====================================================================

        public void createRptCtlFileXML() // write the report control file
        {
            // if this is not a "standard report" the "rename" file name will be set to the reportname from the toc header
            // build the date time constant for the report name
            svDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // get the Node Description
            qNDescription = "";
            DoReadNode(svNodeId);

            // Remove characters from the node name that can't be used in a file name
            string fwkNodeName;
            fwkNodeName = qNDescription;
            fwkNodeName = fwkNodeName.Replace(@"\", "");
            fwkNodeName = fwkNodeName.Replace("/", "");
            fwkNodeName = fwkNodeName.Replace("<", "");
            fwkNodeName = fwkNodeName.Replace(">", "");
            fwkNodeName = fwkNodeName.Replace(":", "");
            fwkNodeName = fwkNodeName.Replace("*", "");
            fwkNodeName = fwkNodeName.Replace("?", "");
            fwkNodeName = fwkNodeName.Replace("\"", "");
            svNodeName = fwkNodeName;                        // save folder name for rq entry update

            svRptFileName = svTargetName;
            // set temp folder for report processing. Naming is "Report Name_Folder_timestamp_temp"
            tempDir = svTempDir + svNodeId + "_" + DateTime.Now.ToString("HHmmss.fffff") + @"\";

            // build the control file name
            svCtlFileName = svTargetDirectory + "P" + svRQId9 + "_PDFCTL.xml"; // put the file name together
            CreateRptCtlFileXML_381(); // yes
        }

        // create builds array and uses Aspose to build Final Report file.
        public void CreateRptCtlFileXML_381()
        {
            string fileExt, secLabel, secLevel;
            var pdfList = new string[rpct + 1, 4];   // set array for processing order of reports.

            // set aspose license
            Setlicense();
            if (!Directory.Exists(tempDir))
            {
                // Attempt to create the temp directory for report processing
                DirectoryInfo Di = null;
                try
                {
                    Di = Directory.CreateDirectory(tempDir);
                }
                catch (Exception ex) // If it fails, handle the error condition
                {
                    // Log Failure
                    LogMsg("ERROR: There was an error creating a temp directory for report processing in folder \"" + tempDir.ToString() + "\". The exception that was thrown is: " + ex.Message, false);
                    // Attempt to clean directory
                    try
                    {
                        Di.Delete();
                    }
                    catch
                    {
                    }
                }
                // TODO: Handle error condition
            }

            // loop through the rptPiecesTable, writing one line for each entry
            // Each entry rptPiecesTable Consists of:
            // 1. filename
            // 2. label
            // 3. level
            // 4. file path
            // 5. section name

            // clear out pdflist array
            Array.Clear(pdfList, 0, pdfList.Length);
            var loopTo = rpct;
            for (rpdx = 1; rpdx <= loopTo; rpdx++)
            {
                if (!string.IsNullOrEmpty(rptPiecesTable[rpdx, 1]))
                {
                    wkFileName = rptPiecesTable[rpdx, 1];
                    wkFileLabel = HttpUtility.HtmlEncode(rptPiecesTable[rpdx, 2]);
                    wkFileLabel = wkFileLabel.Replace("'", "&#39;");
                    pdfList[rpdx, 0] = wkFileLabel;
                }
                else
                {
                    wkFileName = "HEADER";
                }

                secLevel = HttpUtility.HtmlEncode(rptPiecesTable[rpdx, 3]);
                wkFileDir = HttpUtility.HtmlEncode(rptPiecesTable[rpdx, 4]);
                wkFileDir = wkFileDir.Replace("'", "&#39;");
                secLabel = HttpUtility.HtmlEncode(rptPiecesTable[rpdx, 5]);
                fileExt = Path.GetExtension(wkFileName).ToLower();

                // convert files to pdf
                string temp = "";
                if (wkFileName != "HEADER")
                {
                    temp = Asposetest(wkFileName, wkFileDir, fileExt);
                }
                else
                {
                    temp = "";
                }

                pdfList[rpdx, 1] = temp;
                pdfList[rpdx, 2] = secLabel;
                pdfList[rpdx, 3] = secLevel;
            }

            // Places TOC info info from DB into variables 
            ReadTOC();
            try
            {
                // create final pdf
                CreateFinalPdf(svRptFileName, svTargetDirectory, pdfList);
            }
            catch (Exception ex)
            {
                // set report as finish errors occured in report queue. 
                string error = ex.ToString();
                QryRptFinish(svRQId, "E", svTOCReportName, ref error);
                objEventLog.WriteEntry("Report Queue Exception Error : " + error, EventLogEntryType.Error);
            } // Error message in Event log
        }

        public void CreateFinalPdf(string filename, string rptLoc, string[,] fileEntries)
        {

            // create PdfFileEditor object
            var pdfEditor = new PdfFileEditor();
            string filenameTmp = "";
            string tempRpt = "";
            if (filename.Length > 30)
            {
                filenameTmp = filename.Substring(0, 30);
                tempRpt = Path.GetFileNameWithoutExtension(filenameTmp) + "- temp.pdf";
            }
            else
            {
                filenameTmp = filename;
                tempRpt = Path.GetFileNameWithoutExtension(filenameTmp) + "- temp.pdf";
            }

            var templist = new string[fileEntries.GetUpperBound(0) + 1];
            string cvrPage = null;
            var textStamp = new Aspose.Pdf.TextStamp("");

            int cnt = 0;

            for (int I = 0, loopTo = fileEntries.GetUpperBound(0); I <= loopTo; I++)
            {
                if (fileEntries[I, 0] is object)
                {
                    templist[cnt] = fileEntries[I, 1];   // make string array to concatonate all pdfs created.
                    var info = new PdfFileInfo(fileEntries[I, 1]);    // get pdf information
                    cnt += 1;
                }
            }
            templist = templist.Where(p => p != null).ToArray();
            if (false == pdfEditor.Concatenate(templist, tempDir + tempRpt))
            {
                svErrorInd = "C";
                svErrorMsg = "** Error processing merged PDF input File **";
                objEventLog.WriteEntry("Final " + svNodeId + " PDF Output : " + pdfEditor.LastException.Message + Environment.NewLine + pdfEditor.LastException.StackTrace, EventLogEntryType.Error); // Error message in Event log
            }

            var pdfdoc1 = new Aspose.Pdf.Document(tempDir + tempRpt);
            int counter = 1;
            // create page number stamp
            if (svFooterYN == "Y")
            {
                string pageNum;

                // set font family for TOC
                if (svFooterFont == "H")
                {
                    svFooterFont = "Helvetica";
                }
                else
                {
                    svFooterFont = "Times";
                }

                // set standard page numbering format
                if (svStdPageNumbers == "Y")
                {
                    var pageNumberStamp = new Aspose.Pdf.PageNumberStamp();
                    pageNumberStamp.Background = false;
                    pageNumberStamp.Format = "Page # of " + pdfdoc1.Pages.Count;
                    pageNumberStamp.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootBMargin));
                    pageNumberStamp.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootLMargin));
                    pageNumberStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Left;
                    pageNumberStamp.StartingNumber = 1;

                    // set text properties
                    pageNumberStamp.TextState.Font = FontRepository.FindFont(svFooterFont);
                    pageNumberStamp.TextState.FontSize = Convert.ToInt32(svFooterFontSize);
                    pageNumberStamp.TextState.ForegroundColor = Aspose.Pdf.Color.Black;

                    // set time stamp properties
                    var currentTime = DateTime.Now;  // get current time and make a string from it
                    var DateTimeStamp = new Aspose.Pdf.PageNumberStamp();
                    DateTimeStamp.Background = false;
                    DateTimeStamp.Format = Convert.ToString(currentTime);
                    DateTimeStamp.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootBMargin));
                    DateTimeStamp.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootRMargin));
                    DateTimeStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Right;
                    // set text properties
                    DateTimeStamp.TextState.Font = FontRepository.FindFont(svFooterFont);
                    DateTimeStamp.TextState.FontSize = Convert.ToInt32(svFooterFontSize);
                    DateTimeStamp.TextState.ForegroundColor = Aspose.Pdf.Color.Black;
                    foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
                    {
                        // add page number stamps
                        pdfdoc1.Pages[counter].AddStamp(pageNumberStamp);
                        // add date time stamp
                        pdfdoc1.Pages[counter].AddStamp(DateTimeStamp);
                        counter += 1;
                    }
                }
                else
                {

                    // whether the stamp is background
                    if (!string.IsNullOrEmpty(svFooterOptL))
                    {
                        var pageStampLeft = new Aspose.Pdf.PageNumberStamp();
                        string currentDate;
                        counter = 1;

                        // configure stamp based on options selected
                        pageStampLeft.Background = false;
                        switch (svFooterOptL ?? "")
                        {
                            case "D":
                                {
                                    currentDate = DateTime.Now.ToString("MM/dd/yyyy");
                                    pageStampLeft.Format = currentDate;
                                    break;
                                }

                            case "DT":
                                {
                                    currentDate = Convert.ToString(DateTime.Now);
                                    pageStampLeft.Format = currentDate;
                                    break;
                                }

                            case "T":
                                {
                                    pageStampLeft.Format = svFooterL;
                                    break;
                                }

                            case "P":
                                {
                                    pageNum = "Page # of " + pdfdoc1.Pages.Count;
                                    pageStampLeft.Format = pageNum;
                                    break;
                                }
                        }

                        pageStampLeft.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootBMargin));
                        pageStampLeft.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootLMargin));
                        pageStampLeft.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootRMargin));
                        pageStampLeft.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Left;
                        pageStampLeft.StartingNumber = 1;

                        // set text properties
                        pageStampLeft.TextState.Font = FontRepository.FindFont(svFooterFont);
                        pageStampLeft.TextState.FontSize = Convert.ToInt32(svFooterFontSize);
                        pageStampLeft.TextState.ForegroundColor = Aspose.Pdf.Color.Black;

                        // add stamp to each page in Document
                        foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
                        {
                            pdfdoc1.Pages[counter].AddStamp(pageStampLeft);
                            counter += 1;
                        }
                    }

                    if (!string.IsNullOrEmpty(svFooterOptR))
                    {
                        var pageStampRight = new Aspose.Pdf.PageNumberStamp();
                        string currentDate;
                        counter = 1;

                        // configure stamp based on options selected
                        pageStampRight.Background = false;
                        switch (svFooterOptR ?? "")
                        {
                            case "D":
                                {
                                    currentDate = DateTime.Now.ToString("MM/dd/yyyy");
                                    pageStampRight.Format = currentDate;
                                    break;
                                }

                            case "DT":
                                {
                                    currentDate = Convert.ToString(DateTime.Now);
                                    pageStampRight.Format = currentDate;
                                    break;
                                }

                            case "T":
                                {
                                    pageStampRight.Format = svFooterR;
                                    break;
                                }

                            case "P":
                                {
                                    pageNum = "Page # of " + pdfdoc1.Pages.Count;
                                    pageStampRight.Format = pageNum;
                                    break;
                                }
                        }

                        pageStampRight.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootBMargin));
                        pageStampRight.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootLMargin));
                        pageStampRight.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootRMargin));
                        pageStampRight.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Right;
                        pageStampRight.StartingNumber = 1;

                        // set text properties
                        pageStampRight.TextState.Font = FontRepository.FindFont(svFooterFont);
                        pageStampRight.TextState.FontSize = Convert.ToInt32(svFooterFontSize);
                        pageStampRight.TextState.ForegroundColor = Aspose.Pdf.Color.Black;

                        // add stamp to each page in Document
                        foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
                        {
                            pdfdoc1.Pages[counter].AddStamp(pageStampRight);
                            counter += 1;
                        }
                    }

                    if (!string.IsNullOrEmpty(svFooterOptC))
                    {
                        var pageStampCenter = new Aspose.Pdf.PageNumberStamp();
                        string currentDate;
                        counter = 1;

                        // configure stamp based on options selected
                        pageStampCenter.Background = false;
                        switch (svFooterOptC ?? "")
                        {
                            case "D":
                                {
                                    currentDate = DateTime.Now.ToString("MM/dd/yyyy");
                                    pageStampCenter.Format = currentDate;
                                    break;
                                }

                            case "DT":
                                {
                                    currentDate = Convert.ToString(DateTime.Now);
                                    pageStampCenter.Format = currentDate;
                                    break;
                                }

                            case "T":
                                {
                                    pageStampCenter.Format = svFooterC;
                                    break;
                                }

                            case "P":
                                {
                                    pageNum = "Page # of " + pdfdoc1.Pages.Count;
                                    pageStampCenter.Format = pageNum;
                                    break;
                                }
                        }

                        pageStampCenter.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootBMargin));
                        pageStampCenter.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootLMargin));
                        pageStampCenter.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svFootRMargin));
                        pageStampCenter.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                        pageStampCenter.StartingNumber = 1;

                        // set text properties
                        pageStampCenter.TextState.Font = FontRepository.FindFont(svFooterFont);
                        pageStampCenter.TextState.FontSize = Convert.ToInt32(svFooterFontSize);
                        pageStampCenter.TextState.ForegroundColor = Aspose.Pdf.Color.Black;

                        // add stamp to each page in Document
                        foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
                        {
                            pdfdoc1.Pages[counter].AddStamp(pageStampCenter);
                            counter += 1;
                        }
                    }
                }

                pdfdoc1.Save(tempDir + tempRpt);
            }

            svWaterYN = 'N'; // set to N until ready to be impliemented with configurable options
            if (Convert.ToString(svWaterYN) == "Y")  // Not Implimented yet
            {
                // create text stamp
                textStamp = new Aspose.Pdf.TextStamp("Sample Watermark");
                // set whether stamp is background															   
                textStamp.Background = true;
                // set origin
                textStamp.XIndent = 100;
                textStamp.YIndent = 100;
                // rotate stamp
                textStamp.RotateAngle = 45.0f;
                // set text properties
                textStamp.TextState.Font = FontRepository.FindFont("Verdana");
                textStamp.TextState.FontSize = 40.0f;
                textStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                textStamp.VerticalAlignment = Aspose.Pdf.VerticalAlignment.Center;
                textStamp.TextState.ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Gray);
                textStamp.Opacity = 0.75d;
                counter = 1;  // reset counter to 1
                foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
                {
                    // Add watermark
                    pdfdoc1.Pages[counter].AddStamp(textStamp);
                    counter += 1;
                }
            }

            // ***working code just commented out until ready to be implimented with configurable options
            // svHeaderYN = CChar("N")

            // If svHeaderYN = "Y" Then  ' Not Implimented yet
            // 'create header
            // imageStamp = New Aspose.Pdf.ImageStamp("aspose-logo.jpg")
            // 'set properties of the stamp
            // imageStamp.TopMargin = 10
            // imageStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center
            // imageStamp.VerticalAlignment = Aspose.Pdf.VerticalAlignment.Top
            // counter = 1	 'reset counter to 1
            // For Each Page As Aspose.Pdf.Page In pdfdoc1.Pages
            // 'Add header image
            // 'pdfdoc1.Pages(counter).AddStamp(imageStamp)
            // counter += 1
            // Next
            // End If

            int pgcount = 1;
            string tempToc = "";

            // loop through pageIndexes to add book marks
            for (int I = 0, loopTo1 = fileEntries.GetUpperBound(0); I <= loopTo1; I++)
            {
                if (fileEntries[I, 0] is object)
                {
                    string docLbl = HttpUtility.HtmlDecode(fileEntries[I, 0]);
                    string docName = fileEntries[I, 1];
                    var info = new PdfFileInfo(docName);  // get page count of current pdf
                                                          // add bookmarks
                    var pdfOutline = new Aspose.Pdf.OutlineItemCollection(pdfdoc1.Outlines);
                    pdfOutline.Title = docLbl;   // Sets Bookmark name as DOC Label
                                                 // Set the destination page number
                    pdfOutline.Action = new Aspose.Pdf.Annotations.GoToAction(pdfdoc1.Pages[pgcount]);
                    // Add bookmark in the document's outline collection.
                    pdfdoc1.Outlines.Add(pdfOutline);

                    // add page count of Doc to total page count	
                    pgcount += info.NumberOfPages;
                }
                else if (!string.IsNullOrEmpty(fileEntries[I, 2]))
                {
                    string sectionLbl = fileEntries[I, 2];
                    var pdfOutline = new Aspose.Pdf.OutlineItemCollection(pdfdoc1.Outlines);
                    pdfOutline.Title = sectionLbl;   // Sets Bookmark name as DOC Label
                    pdfOutline.Bold = true;
                    // Add bookmark in the document's outline collection.
                    pdfdoc1.Outlines.Add(pdfOutline);
                }
            }

            // checks image area and sets to landscape if necessary.
            foreach (Aspose.Pdf.Page Page in pdfdoc1.Pages)
            {
                // set page to landscape is necessary.
                Aspose.Pdf.Rectangle rect = Page.MediaBox;
                if (Page.Rect.Width > Page.Rect.Height)
                {
                    Page.PageInfo.IsLandscape = true;
                }
                else
                {
                    Page.PageInfo.IsLandscape = false;
                }
            }

            // Builds TOC for fiinal report.

            if (Convert.ToString(svTocYN) == "Y")
            {
                tempToc = ProcTOC(fileEntries, tempDir); // processes the TOC to get the exact count pages it spans.
                pdfdoc1.Save(tempDir + "TOC-temp.pdf");
                pdfEditor.Concatenate(tempToc, tempDir + "TOC-temp.pdf", tempDir + "tocFinal-temp.pdf"); // merge TOC and existing report pdf
                pdfdoc1 = new Aspose.Pdf.Document(tempDir + "tocFinal-temp.pdf");  // add combined file into PDF object
            }

            filename = HttpUtility.HtmlDecode(filename);

            String finalFname = rptLoc + "/" + filename;
            String finalDir = Path.GetDirectoryName(finalFname);
            
            if(!Directory.Exists(finalDir))
            {
                Directory.CreateDirectory(finalDir);
            }

            if (Convert.ToString(svCoverYN) == "Y")
            {
                // save pdf to final location and report name
                pdfdoc1.Save(tempDir + "Final-temp.pdf");

                // find and coverts cover page to pdf
                cvrPage = ProcessCover();

                // add Cover page and save final report
                if (false == pdfEditor.Concatenate(cvrPage, tempDir + "Final-temp.pdf", finalFname))
                {
                    svErrorInd = "C";
                    svErrorMsg = "** Error Processing Final PDF Output File **";
                    objEventLog.WriteEntry("Final " + svNodeId + "PDF Output : " + pdfEditor.LastException.Message + Environment.NewLine + pdfEditor.LastException.StackTrace, EventLogEntryType.Error); // Error message in Event log
                }
            }
            // TODO: Find out what went wrong
            else
            {
                // no cover page save final report
                pdfdoc1.Save(finalFname);
            }

            try
            {
                // delete the temp folder user for report processing.
                Directory.Delete(tempDir, true);
            }
            catch (Exception ex)
            {
                objEventLog.WriteEntry("Report Queue Exception Error : " + ex.ToString(), EventLogEntryType.Error);
            } // Error message in Event log

            if (svErrorInd == "Y")
            {
                // adds final report to document table
                int argnodeid = svNodeId;
                AddRptDocument(ref argnodeid, svTOCReportName, filename);
                // set report as finish errors occured in report queue. 
                QryRptFinish(svRQId, "E", svTOCReportName, ref svErrorMsg);
            }
            else if (svErrorInd == "C") // File does not exist, don't add doc
            {
                // set report as finish errors occured in report queue. 
                QryRptFinish(svRQId, "E", svTOCReportName, ref svErrorMsg);
            }
            else
            {
                // adds final report to document table
                int argnodeid1 = svNodeId;
                AddRptDocument(ref argnodeid1, svTOCReportName, filename);
                string argrptError = "";
                // set report as finshed  no errors in report queue. 
                QryRptFinish(svRQId, "C", svTOCReportName, rptError: ref argrptError);
            }
        }

        public string ProcTOC(string[,] fileEntries, string tempFolder)
        {
            var doc = new Aspose.Words.Document();
            var builder = new Aspose.Words.DocumentBuilder(doc);
            double tabstop;
            double pgWidth;
            double pgMargins = Convert.ToDouble(svTocLeftMargin) + Convert.ToDouble(svTocRightMargin);
            int pgcount = 1;

            // set tab stop location and page orientation	 also calculates page size and subtracts
            // a half inch to set the tabstop for the page number in the toc.
            if (svTocLayout == "L")
            {
                builder.PageSetup.Orientation = Aspose.Words.Orientation.Landscape;
                pgWidth = 11d - pgMargins;
                tabstop = pgWidth - 0.05d;
            }
            else
            {
                builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;
                pgWidth = 8.5d - pgMargins;
                tabstop = pgWidth - 0.05d;
            }

            // set font type base on selecting in report screen
            if (svTocFont == "H")
            {
                svTocFont = "Helvetica";
            }
            else
            {
                svTocFont = "Times";
            }

            // add header entry for TOC
            builder.ParagraphFormat.TabStops.Clear();
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.MoveToHeaderFooter(Aspose.Words.HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = Aspose.Words.ParagraphAlignment.Center;
            builder.Font.Bold = true;
            builder.Font.Size = 24;
            builder.Font.Name = svTocFont;
            builder.Write("Table of Contents");
            builder.MoveToDocumentStart();
            builder.PageSetup.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svTocLeftMargin));
            builder.PageSetup.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(Convert.ToDouble(svTocRightMargin));
            builder.PageSetup.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(0.5d);
            builder.PageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(0.5d);
            builder.ListFormat.List = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberUppercaseRomanDot);
            builder.ParagraphFormat.Alignment = Aspose.Words.ParagraphAlignment.Left;

            // loop through fileEntries to add TOC Links						
            for (int I = 0, loopTo = fileEntries.GetUpperBound(0); I <= loopTo; I++)
            {
                if (fileEntries[I, 0] is object)  // if document label
                {
                    string docLbl = HttpUtility.HtmlDecode(fileEntries[I, 0]);
                    string docName = fileEntries[I, 1];
                    var info = new PdfFileInfo(docName);  // get pdf information

                    // Set style for list of docs under heading
                    builder.Font.Name = svTocFont;
                    builder.Font.Bold = false;
                    builder.Font.Size = Convert.ToDouble(svTocFontSize);
                    builder.ListFormat.ListLevelNumber = 1;
                    builder.ListFormat.ListLevel.NumberPosition = Aspose.Words.ConvertUtil.InchToPoint(0.5d); // sets indent to .5 inch
                    builder.ListFormat.ListLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Left;
                    builder.ListFormat.ListLevel.TrailingCharacter = Aspose.Words.Lists.ListTrailingCharacter.Space;
                    builder.ListFormat.ListLevel.NumberStyle = Aspose.Words.NumberStyle.None;
                    builder.ListFormat.ListLevel.NumberFormat = ""; // removes trailing period in outline numbers 

                    // clears tab stop and adds in 1 with lead character as dots.
                    builder.ParagraphFormat.TabStops.Clear();
                    builder.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.InchToPoint(tabstop), Aspose.Words.TabAlignment.Right, Aspose.Words.TabLeader.Dots);
                    builder.Write(docLbl + "\t");  // writes label and enters in tab

                    // if last item in list, use write instead of write line to avoids extra line in TOC.
                    if (I == fileEntries.GetUpperBound(0))
                    {
                        builder.Write(pgcount.ToString());
                    }
                    else
                    {
                        builder.Writeln(pgcount.ToString());
                    }

                    pgcount += info.NumberOfPages;
                }
                else if (!string.IsNullOrEmpty(fileEntries[I, 2]))    // if section label
                {
                    string rptSection = fileEntries[I, 2];
                    builder.ParagraphFormat.TabStops.Clear();
                    // sets formatting for section header.
                    builder.ListFormat.ListLevel.NumberPosition = 0;  // makes outList setion header start where margin is set.
                    builder.ListFormat.ListLevel.Alignment = Aspose.Words.Lists.ListLevelAlignment.Left;
                    builder.ListFormat.ListLevel.TrailingCharacter = Aspose.Words.Lists.ListTrailingCharacter.Space;
                    builder.ListFormat.ListLevel.NumberStyle = Aspose.Words.NumberStyle.None;
                    builder.ListFormat.ListLevel.NumberFormat = ""; // removes trailing period in outline numbers 
                    builder.Font.Size = Convert.ToDouble(svTocFontSize);
                    builder.Font.Name = svTocFont;
                    builder.Font.Bold = true;
                    builder.ListFormat.ListLevelNumber = 0;

                    // if last item in list, use write instead of write line to avoids extra line in TOC.
                    if (I == fileEntries.GetUpperBound(0))
                    {
                        builder.Write(rptSection);
                    }
                    else
                    {
                        builder.Writeln(rptSection);
                    }
                }
            }

            doc.Save(tempDir + "TOC.pdf", Aspose.Words.SaveFormat.Pdf); // save PDF of TOC	
            return tempDir + "TOC.pdf";
        }

        public void QryRptFinish(int qRQID, string qStatus, string rptLabel, [Optional, DefaultParameterValue("")] ref string rptError)
        {
            SUSPLAN_REPORT_Q reportQ = _context.SUSPLAN_REPORT_Q.Find(qRQID);
            if (reportQ != null)
            {
                reportQ.FINISHED_DATE = DateTime.Now;
                reportQ.STATUS = qStatus;
                reportQ.ERROR_MESSAGE = rptError.Substring(0, Math.Min(rptError.Length, 499)) ?? null;
                _context.Entry(reportQ).State = EntityState.Modified;
                _context.SaveChanges();
            }
        }

        public string ProcessCover()
        {
            string fwkFileName;
            string fwkFilePath;
            string fwkFullFN;
            string fileExt;
            string tempName;
            string tempCover;
            string temppath;
            string[] temp;
            var listDocument = _context.SUSPLAN_DOCUMENTS.Where(d => d.DOCUMENT_ID == svCvrPageID).ToList();
            int docFolder;
            string docLabel = "";

            foreach (var item in listDocument)
            {
                // gets Document label of Cover Doc
                docLabel = string.IsNullOrEmpty(item.DOCUMENT_LABEL) ? "MISSING" : item.DOCUMENT_LABEL;
                docFolder = item.NODE_ID ?? 0;
            }

            // build the full path to the document
            svDocDirId = ""; // clear the docdirectory id
            temppath = ProcessCvrDoc(docLabel);
            temp = temppath.Split(',');

            // splits out path and extension of cover page found
            fwkFilePath = temp[0];
            fileExt = temp[1];
            fwkFileName = temp[2];
            // fwkFullFN = tempDir & fwkFileName ' build the full file path/name

            // convert cover document to PDF
            tempCover = Asposetest(fwkFileName, fwkFilePath, fileExt);
            tempName = Path.GetFileNameWithoutExtension(fwkFileName);

            // returns path of converted cover page.
            fwkFullFN = tempDir + tempName + ".pdf";
            return tempCover;
        }

        public string ProcessCvrDoc(string fwkTocDocLabel)
        {
            string fwkFileName = "";
            string fwkFilePath = "";
            string fwkFileType;
            string fwkFullFN;
            string fwkCascades;
            DateTime? fwkNFileCreateDate;
            DateTime? fwkNFileDate;
            DateTime? fwkNFileUpdateDate;
            DateTime? fwkFileDate;
            string fwkNFileName;
            var cvrPageDT = new System.Data.DataTable();
            var cvrPageDS = new System.Data.DataSet();
            rc = 0; // initialize return code

            // look for a document record which matches the TocDoc Label
            // search in each node up the chain
            // if there are multiple documents with the label within a given node, use the one with
            // the latest create or update date

            fwkFileDate = DateTime.Parse("1900-01-01");
            // clear the node table
            for (ntdx = 0; ntdx <= 10; ntdx++)
            {
                nodeTbl[ntdx] = 0;
                Set_nodeDocDirTbl(ntdx, "");
            }

            LoadNodeTbl(svNodeId); // loads the Folder ID array
            ntdx = 1;
            for (ntdx = 1; ntdx <= 10; ntdx++)
            {
                if (ntdx > nodeTbl[0]) // past the end of the node entries?
                {
                    break; // yep
                }

                qNodeId = nodeTbl[ntdx]; // get the node id
                qDocLabel = fwkTocDocLabel; // get the document label 
                var listDocument = _context.SUSPLAN_DOCUMENTS.Where(d => d.NODE_ID == qNodeId && d.DOCUMENT_LABEL.ToUpper() == qDocLabel).ToList();
                foreach (var item in listDocument)
                {
                    fwkNFileName = item.PHYSICAL_NAME;
                    fwkNFileCreateDate = item.CREATE_DATE;
                    fwkNFileUpdateDate = item.UPDATE_DATE;
                    fwkCascades = item.DOC_CASCADES_YN;
                    if (ntdx > 1 & fwkCascades == "N") // if doc is above this node and does not cascade
                    {
                        fwkNFileName = ""; // don't use it
                    }

                    fwkNFileDate = fwkNFileCreateDate; // get the file create date
                    if (fwkNFileUpdateDate.HasValue) // is there an update date?
                    {
                        fwkNFileDate = fwkNFileUpdateDate; // yep - use it
                    }

                    if (!string.IsNullOrEmpty(fwkNFileName)) // got a file to use?
                    {
                        if (string.IsNullOrEmpty(fwkFileName)) // yep - is this the first match?
                        {
                            fwkFileName = fwkNFileName; // yep - store the file name
                            fwkFileDate = fwkNFileDate; // and date
                        }
                        else if (fwkNFileDate > fwkFileDate)  // does this one have a later update date?
                        {
                            fwkFileName = fwkNFileName; // yep - store the file name
                            fwkFileDate = fwkNFileDate; // and date
                        }
                    }
                }

                if (!string.IsNullOrEmpty(fwkFileName))
                    break; // if file was found stop looking
            }

            if (!string.IsNullOrEmpty(fwkFileName))
            {
                // build the full path to the document
                fwkFilePath = ""; // initialize the result field
                svDocDirId = ""; // clear the docdirectory id
                while (ntdx <= nodeTbl[0]) // build the path from this node up
                {
                    fwkFilePath = nodeTbl[ntdx].ToString() + @"\" + fwkFilePath;
                    if (string.IsNullOrEmpty(svDocDirId) & !string.IsNullOrEmpty(Get_nodeDocDirTbl(ntdx))) // is this the lowest docdirectory id?
                    {
                        svDocDirId = Get_nodeDocDirTbl(ntdx); // yep - store it
                    }

                    ntdx = ntdx + 1;
                }

                fwkFilePath = documentPath + @"\" + fwkFilePath; // add the base path

                // check whether the file actually exists
                string fileExt = Path.GetExtension(fwkFileName);
                fwkFullFN = fwkFilePath + fwkFileName; // build the full file path/name
                if (File.Exists(fwkFullFN) == true)
                {
                    fwkFileType = fwkFileName.Substring(fwkFileName.Length - 4, 4).ToUpper(); // get the file extention
                    if (fwkFileType == ".RPT") // Crystal Report?
                    {
                        svCrystalInd = "Y"; // yes - set indicator
                    }
                }

                return fwkFilePath + "," + fileExt + "," + fwkFileName;
            }
            else // No physical file was found uploaded
            {
                return fwkFilePath + ",.err,ERROR";
            }
        }

        public void AddRptDocument(ref int nodeid, string reportLabel, string fileName)
        {
            SUSPLAN_DOCUMENTS newDocument = new SUSPLAN_DOCUMENTS()
            {
                NODE_ID = nodeid,
                DOCUMENT_LABEL = reportLabel,
                PHYSICAL_NAME = fileName,
                STATUS = "N",
                CREATE_DATE = DateTime.Now,
                UPDATE_DATE = DateTime.Now
            };
            _context.SUSPLAN_DOCUMENTS.Add(newDocument);
            _context.SaveChanges();
        }

        public void ReadTOC() // read the toc header
        {
            var tocHeader = _context.SUSPLAN_TOC_HEADER.Find(svTOCId);
            if (tocHeader == null)
            {
                rc = 1; // set return code
            }
            else
            {
                svTOCNodeId = tocHeader.NODE_ID ?? 0;
                svTOCReportName = tocHeader.REPORT_NAME;

                // remove invalid characters from the report name
                // These characters \/:*?<>| are not valid for a windows file name
                svTOCReportName = svTOCReportName.Replace(@"\", "");
                svTOCReportName = svTOCReportName.Replace("/", "");
                svTOCReportName = svTOCReportName.Replace(":", "");
                svTOCReportName = svTOCReportName.Replace("*", "");
                svTOCReportName = svTOCReportName.Replace("?", "");
                svTOCReportName = svTOCReportName.Replace("<", "");
                svTOCReportName = svTOCReportName.Replace(">", "");
                svTOCReportName = svTOCReportName.Replace("|", "");
                svTOCReportName = svTOCReportName.Replace("&#58;", "");
                if (string.IsNullOrEmpty(svTOCReportName)) // was the name all special characters?
                {
                    svTOCReportName = "no name";
                }

                svTOCReportName = HttpUtility.HtmlEncode(svTOCReportName);
                svTOCPage = tocHeader.TOC_PAGE ?? 0;
                svTocLayout = string.IsNullOrEmpty(tocHeader.TOC_LAYOUT) ? "P" : tocHeader.TOC_LAYOUT;
                svTocLeftMargin = string.IsNullOrEmpty(tocHeader.TOC_LEFT_MARGIN) ? ".25" : tocHeader.TOC_LEFT_MARGIN;
                svTocRightMargin = string.IsNullOrEmpty(tocHeader.TOC_RIGHT_MARGIN) ? ".25" : tocHeader.TOC_RIGHT_MARGIN;
                svTocFont = string.IsNullOrEmpty(tocHeader.TOC_FONT) ? "T" : tocHeader.TOC_FONT;
                svTocFontSize = string.IsNullOrEmpty(tocHeader.TOC_FONT_SIZE) ? "10" : tocHeader.TOC_FONT_SIZE;
                svHeaderStartOnDoc = tocHeader.HEADER_START_ON_DOC ?? 0;
                svHeaderFont = string.IsNullOrEmpty(tocHeader.HEADER_FONT) ? "T" : tocHeader.HEADER_FONT;
                svHeaderFontSize = string.IsNullOrEmpty(tocHeader.HEADER_FONT_SIZE) ? "10" : tocHeader.HEADER_FONT_SIZE;
                svFooterFont = string.IsNullOrEmpty(tocHeader.FOOTER_FONT) ? "T" : tocHeader.FOOTER_FONT;
                svFooterFontSize = string.IsNullOrEmpty(tocHeader.FOOTER_FONT_SIZE) ? "10" : tocHeader.FOOTER_FONT_SIZE;
                svHeaderText = tocHeader.HEADER_TEXT;
                svFooterStartOnDoc = tocHeader.FOOTER_START_ON_DOC ?? 0;
                svFooterL = tocHeader.FOOTER_TEXT_LEFT;
                svFooterC = tocHeader.FOOTER_TEXT_CENTER;
                svFooterR = tocHeader.FOOTER_TEXT_RIGHT;
                svStdPageNumbers = string.IsNullOrEmpty(tocHeader.STD_PAGE_NUMBERS) ? "Y" : tocHeader.STD_PAGE_NUMBERS;
                svTOCFirstSectionId = tocHeader.FIRST_SECTION_ID ?? 0;
                svCoverYN = string.IsNullOrEmpty(tocHeader.COVER_PAGE_YN) ? 'Y' : Convert.ToChar(tocHeader.COVER_PAGE_YN);
                svCvrPageID = tocHeader.COVER_PAGE_ID ?? 0;
                svCvrPage = tocHeader.COVER_PAGE;
                svTocYN = string.IsNullOrEmpty(tocHeader.TOC_YN) ? 'N' : Convert.ToChar(tocHeader.TOC_YN);
                svFootLMargin = string.IsNullOrEmpty(tocHeader.FOOTER_MARGIN_L) ? ".5" : tocHeader.FOOTER_MARGIN_L;
                svFootRMargin = string.IsNullOrEmpty(tocHeader.FOOTER_MARGIN_R) ? ".5" : tocHeader.FOOTER_MARGIN_R;
                svFootBMargin = string.IsNullOrEmpty(tocHeader.FOOTER_MARGIN_B) ? ".5" : tocHeader.FOOTER_MARGIN_B;
                svFooterYN = string.IsNullOrEmpty(tocHeader.FOOTER_YN) ? "N" : tocHeader.FOOTER_YN;
                svFooterOptL = tocHeader.FOOTER_OPT_L;
                svFooterOptR = tocHeader.FOOTER_OPT_R;
                svFooterOptC = tocHeader.FOOTER_OPT_C;
                svFooterL = HttpUtility.HtmlDecode(svFooterL);
                svFooterR = HttpUtility.HtmlDecode(svFooterR);
                svFooterC = HttpUtility.HtmlDecode(svFooterC);
                svTOCReportName = HttpUtility.HtmlDecode(svTOCReportName);
            }
        }

        // ============================================================================= 
        // Read a node record, using the Node ID
        // ============================================================================= 

        public int DoReadNode(int fwkNodeId)
        {
            int doReadNodeRet = default;
            doReadNodeRet = 0; // initialize return code
            var nodes = _context.SUSPLAN_NODES.Find(fwkNodeId);
            if (nodes == null)
            {
                doReadNodeRet = 1; // set return code
            }
            else
            {
                qNParentId = nodes.PARENT_ID ?? 0;
                qPreviousNodeId = nodes.PREVIOUS_NODE_ID ?? 0;
                qNextNodeId = nodes.NEXT_NODE_ID ?? 0;
                qNFirstChildId = nodes.FIRST_CHILD_ID ?? 0;
                qLevelNumber = nodes.LEVEL_NUMBER ?? 0;
                qNDescription = nodes.DESCRIPTION;
                qDocDirId = nodes.DOCDIRECTORY_ID;
            }

            return doReadNodeRet;
        }

        public void ProcessSection(int fwkSectionId, int fwkLevelNo)
        {
            int fwkFirstChildId;
            string fwkSectionName;
            int fwkNextSectionId;
            int fwkFirstTOCDOCId;
            svCrystalInd = "N"; // initialize crystal reports indicator
            svErrorInd = "N"; // initialize error indicator 
            svInputDocInd = "N"; // initialize input document indicator

            // rc = doReadSection(fwkSectionId) ' read the section header record

            // ****Reads section info *******
            rc = 0; // initialize return code
            qSectionId = fwkSectionId;
            var secHeader = _context.SUSPLAN_SEC_HDR.Find(qSectionId);
            if (secHeader == null)
            {
                rc = 1; // set return code
            }
            else
            {
                qSNextSectionId = secHeader.NEXT_SECTION_ID ?? 0;
                qSNFirstChildId = secHeader.CHILD_SECTION_ID ?? 0;
                qSFirstTOCDOCId = secHeader.FIRST_TOCDOC_ID ?? 0;
                qSSectionName = secHeader.SECTION_NAME;
            }

            if (rc != 0) // ID FOUND?
            {
                // Call returnToCaller("Section In Chain Not Found (" & CStr(fwkSectionId) & ") planbuild.aspx")
            }

            fwkNextSectionId = qSNextSectionId; // save data from the section header read
            fwkFirstChildId = qSNFirstChildId;
            fwkFirstTOCDOCId = qSFirstTOCDOCId;
            fwkSectionName = qSSectionName;


            // *************************NEW***************************************
            fwkSecOption = SectionOption(fwkSectionId);  // checks if section should be optional and displayed
                                                         // fwkDocCount = Queries.qrySectionDoc(objDC, fwkSectionId) 'Counts all doc in a section

            if (fwkSecOption == 1)
            {
            }
            // do not add section to rptPiecesTable
            else
            {
                string argfwkTxtFileName = "";
                string argfwkLabel = "";
                int argfwkLevelNo = fwkLevelNo + 1;
                string argfwkFilePath = "";
                AddToRPTable(ref argfwkTxtFileName, ref argfwkLabel, ref argfwkLevelNo, ref argfwkFilePath, ref fwkSectionName);
            } // add an entry to the rptPiecesTable

            // First process the TocDoc chain
            if (fwkFirstTOCDOCId != 0) // is there a TocDoc chain?
            {
                svInputDocInd = "Y"; // indicate at least one input document
                ProcessTocDoc(fwkFirstTOCDOCId, fwkLevelNo); // yes - process it
            }

            // Second, process child sections

            if (fwkFirstChildId != 0) // is there a child section chain?
            {
                ProcessSection(fwkFirstChildId, fwkLevelNo + 1); // yes - process it
            }

            // Third, process the section chain at this level

            if (fwkNextSectionId != 0) // is there a another section in this chain?
            {
                ProcessSection(fwkNextSectionId, fwkLevelNo); // yes - process it
            }
        }

        // ============================================================================= 
        // Read down a tocdoc chain, adding entries to the rptPiecesTable                
        // This routine calls itself recursively to go down  the chain                   
        // ============================================================================= 

        public void ProcessTocDoc(int fwkTocDocId, int fwkLevelNo)
        {
            string fwkTxtFileName;
            string fwkFileName;
            int fwkNextTocDocId;
            string fwkTocDocLabel;
            string fwkFilePath;
            string fwkFileType;
            string fwkFullFN;
            string fwkCascades;
            DateTime? fwkNFileCreateDate;
            DateTime? fwkNFileDate;
            DateTime? fwkNFileUpdateDate;
            DateTime? fwkFileDate;
            string fwkNFileName;
            var fwkOptional = default(int);


            // rc = doReadTocDoc(fwkTocDocId) ' read the tocdoc record
            rc = 0; // initialize return code
            qTocDocId = fwkTocDocId;
            var tocDocument = _context.SUSPLAN_TOC_DOCUMENT.Find(qTocDocId);
            if (tocDocument == null) // ID NOT FOUND
            {
                rc = 1; // set return code
            }
            else
            {
                qNextTocDocId = tocDocument.NEXT_TOCDOC_ID ?? 0;
                qTocDocLabel = tocDocument.DOCUMENT_LABEL;
            }

            if (rc != 0) // ID FOUND?
            {
                // Call returnToCaller("TOC Document In Chain Not Found (" & CStr(fwkTocDocId) & ") planbuild.aspx")
            }

            fwkNextTocDocId = qNextTocDocId; // save data from the tocdoc header read
            fwkTocDocLabel = qTocDocLabel;

            // look for a document record which matches the TocDoc Label
            // search in each node up the chain
            // if there are multiple documents with the label within a given node, use the one with
            // the latest create or update date

            fwkFileName = "";
            fwkFileDate = DateTime.Parse("1900-01-01");
            // clear the node table
            for (ntdx = 0; ntdx <= 10; ntdx++)
            {
                nodeTbl[ntdx] = 0;
                Set_nodeDocDirTbl(ntdx, "");
            }

            LoadNodeTbl(svNodeId); // loads the Folder ID array
            ntdx = 1;
            for (ntdx = 1; ntdx <= 10; ntdx++)
            {
                if (ntdx > nodeTbl[0]) // past the end of the node entries?
                {
                    break; // yep
                }

                qNodeId = nodeTbl[ntdx]; // get the node id
                qDocLabel = fwkTocDocLabel; // get the document label 
                var listDocument = _context.SUSPLAN_DOCUMENTS.Where(d => d.NODE_ID == qNodeId && d.DOCUMENT_LABEL.ToUpper() == qDocLabel).ToList();
                foreach (var item in listDocument)
                {
                    fwkNFileName = item.PHYSICAL_NAME;
                    fwkNFileCreateDate = item.CREATE_DATE;
                    fwkNFileUpdateDate = item.UPDATE_DATE;
                    fwkCascades = string.IsNullOrEmpty(item.DOC_CASCADES_YN) ? "N" : item.DOC_CASCADES_YN;
                    if (ntdx > 1 & fwkCascades == "N") // if doc is above this node and does not cascade
                    {
                        fwkNFileName = ""; // don't use it
                    }

                    fwkNFileDate = fwkNFileCreateDate; // get the file create date
                    if (fwkNFileUpdateDate.HasValue) // is there an update date?
                    {
                        fwkNFileDate = fwkNFileUpdateDate; // yep - use it
                    }

                    if (!string.IsNullOrEmpty(fwkNFileName)) // got a file to use?
                    {
                        if (string.IsNullOrEmpty(fwkFileName)) // yep - is this the first match?
                        {
                            fwkFileName = fwkNFileName; // yep - store the file name
                            fwkFileDate = fwkNFileDate; // and date
                        }
                        else if (fwkNFileDate > fwkFileDate)  // does this one have a later update date?
                        {
                            fwkFileName = fwkNFileName; // yep - store the file name
                            fwkFileDate = fwkNFileDate; // and date
                        }
                    }
                }

                if (!string.IsNullOrEmpty(fwkFileName))
                    break; // if file was found stop looking
            }

            // Check if document is optional
            tocDocument = _context.SUSPLAN_TOC_DOCUMENT.Find(fwkTocDocId);
            if (tocDocument == null)
            {
                rc = 1; // set return code
            }
            else
            {
                fwkOptional = tocDocument.DOC_OPTION ?? 0; // gets Doc Optional Status
            }

            if (string.IsNullOrEmpty(fwkFileName)) // file found?
            {
                if (fwkOptional == 1) // If doc is missing and optional do not add to TOC
                {
                }
                // if document is optional no message is displayed.
                else
                {
                    // Call displayError("A Document Needed For This Report Could Not Be Located. Label = " & fwkTocDocLabel) ' nope - display error message

                    fwkTxtFileName = DoAddErrFile(fwkTocDocId, fwkTocDocLabel); // create a text file in place of the missing document
                    svErrorMsg = "**Document is Missing**";
                    string argfwkLabel = fwkTocDocLabel + " *Missing*";
                    int argfwkLevelNo = fwkLevelNo + 1;
                    string argfwkSectionName = "";
                    AddToRPTable(ref fwkTxtFileName, ref argfwkLabel, ref argfwkLevelNo, ref svTempDir, ref argfwkSectionName);
                }  // add an entry to the rptPiecesTable
            }
            else
            {
                // build the full path to the document
                fwkFilePath = ""; // initialize the result field
                svDocDirId = ""; // clear the docdirectory id
                while (ntdx <= nodeTbl[0]) // build the path from this node up
                {
                    fwkFilePath = nodeTbl[ntdx].ToString() + @"\" + fwkFilePath;
                    if (string.IsNullOrEmpty(svDocDirId) & !string.IsNullOrEmpty(Get_nodeDocDirTbl(ntdx))) // is this the lowest docdirectory id?
                    {
                        svDocDirId = Get_nodeDocDirTbl(ntdx); // yep - store it
                    }

                    ntdx = ntdx + 1;
                }

                // fwkFilePath = svTargetDirectory & "\" & fwkFilePath	' add the base path
                // fwkFilePath = Application("docdirectory") + "\" + fwkFilePath			' add the base path
                fwkFilePath = documentPath + @"\" + fwkFilePath; // add the base path



                // check whether the file actually exists
                fwkFullFN = fwkFilePath + fwkFileName; // build the full file path/name

                // Response.Write("<BR>File name = " + fwkFullFN)

                if (File.Exists(fwkFullFN) == true)
                {
                    int argfwkLevelNo1 = fwkLevelNo + 1;
                    string argfwkSectionName1 = "";
                    AddToRPTable(ref fwkFileName, ref fwkTocDocLabel, ref argfwkLevelNo1, ref fwkFilePath, ref argfwkSectionName1); // add an entry to the rptPiecesTable
                    fwkFileType = fwkFileName.Substring(fwkFileName.Length - 4, 4).ToUpper(); // get the file extention
                    if (fwkFileType == ".RPT") // Crystal Report?
                    {
                        svCrystalInd = "Y"; // yes - set indicator
                    }
                }
                else
                {
                    // Call displayError("A document for this report was not found on the server. The document will need to be uploaded again. The label is " & fwkTocDocLabel) ' nope - display error message
                    fwkTxtFileName = DoAddErrFile(fwkTocDocId, fwkTocDocLabel); // create a text file in place of the missing document
                    svErrorMsg = "**Document is Missing**";
                    string argfwkLabel1 = fwkTocDocLabel + " *Missing*";
                    int argfwkLevelNo2 = fwkLevelNo + 1;
                    string argfwkSectionName2 = "";
                    AddToRPTable(ref fwkTxtFileName, ref argfwkLabel1, ref argfwkLevelNo2, ref svTempDir, ref argfwkSectionName2);
                } // add an entry to the rptPiecesTable
            }

            // process the tocdoc chain 

            if (fwkNextTocDocId != 0) // is there a another tocdoc in this chain?
            {
                ProcessTocDoc(fwkNextTocDocId, fwkLevelNo); // yes - process it
            }
        }


        // ============================================================================= 
        // load the node table by reading up the node tree
        // ============================================================================= 

        public void LoadNodeTbl(int fwkNodeId)
        {

            // Read the Node record
            rc = DoReadNode(fwkNodeId);
            if (rc != 0) // ID FOUND?
            {
                // Call returnToCaller("Node In Chain Not Found (" & CStr(fwkNodeId) & ") planbuild.aspx")

            }

            ntdx = nodeTbl[0] + 1; // increment the count
            nodeTbl[0] = ntdx; // store the updated count
            nodeTbl[ntdx] = fwkNodeId; // store this node id
            Set_nodeDocDirTbl(ntdx, qDocDirId); // store this nodes docdirectory id
            if (qNParentId != 0) // is there a parent?
            {
                LoadNodeTbl(qNParentId); // yep - call this function recursively to add it to the table
            }
        }

        // create a text file as a as a placeholder for a missing document
        public string DoAddErrFile(int fwkTocDocId, string fwkTocDocLabel)
        {
            string doAddErrFileRet = default;
            string fwkPad;
            string fwkFileName;
            string fwkTOCDOCIdStr;
            string fwkTOCDOCIdStr5;
            Setlicense();

            // build the file name
            // format a 5 digit text TOCDOC id
            fwkTOCDOCIdStr = fwkTocDocId.ToString(); // get the TOCDOC ID 
            fwkPad = new string('0', 5 - fwkTOCDOCIdStr.Length); // get any needed padding
            fwkTOCDOCIdStr5 = fwkPad + fwkTOCDOCIdStr; // generate 5 character id
            fwkFileName = "T" + svRQId9 + "_E" + fwkTOCDOCIdStr5 + ".txt"; // put the file name together
            doAddErrFileRet = fwkFileName; // pass the file name back to the caller
            fwkFileName = svTempDir + fwkFileName; // add the path to the file name

            // create blank pdf
            Aspose.Pdf.Document pdf1 = new Aspose.Pdf.Document();//Add blank pdf
            var page = pdf1.Pages.Add();// Add blank page
            string tempdoc = Path.GetFileNameWithoutExtension(fwkFileName);  // get filename without ext.
            pdf1.Save(svTempDir + tempdoc + ".pdf");  // save blank pdf
            fwkFileName = svTempDir + tempdoc + ".pdf";

            // create text stamp
            var MissingStamp = new Aspose.Pdf.TextStamp("**Document is Missing**");
            // set whether stamp is background															   
            MissingStamp.Background = false;
            // set text properties
            MissingStamp.TextState.Font = FontRepository.FindFont("Verdana");
            MissingStamp.TextState.FontSize = 40.0f;
            MissingStamp.TextState.FontStyle = FontStyles.Italic;
            MissingStamp.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;
            MissingStamp.VerticalAlignment = Aspose.Pdf.VerticalAlignment.Center;
            MissingStamp.TextState.ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Black);


            // open blank pdf
            var doc = new Aspose.Pdf.Document(fwkFileName);
            // add watermark
            doc.Pages[1].AddStamp(MissingStamp);
            // save changes
            doc.Save(fwkFileName, Aspose.Pdf.SaveFormat.Pdf);
            return Path.GetFileName(fwkFileName);
        }

        public void Setlicense()
        {
            // path to the license file - set according to your license file location
            //string totalLicense = Path.GetDirectoryName(documentPath);
            string totalLicense = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            totalLicense = totalLicense + @"\lic\ASpose.total.lic.txt";   // look for sustainable_planner directory
            //int index = Environment.CurrentDirectory.IndexOf("ReportProcess");
            //string sp4APIPath = Environment.CurrentDirectory.Substring(0, index) + @"\SP4_API\includes";
            //Directory.CreateDirectory(sp4APIPath);
            //if (File.Exists(totalLicense))
            //{
            //    //copy to the SP4 curren t directory / includes folder(create if not existing).
            //    File.Copy(totalLicense, sp4APIPath + @"\ASpose.total.lic.txt", true);
            //}
            //else
            //{
            //    //check SP4 current directory/includes(create if not existing) folder and change below to the new path.
            //    totalLicense = sp4APIPath + @"\ASpose.total.lic.txt";
            //}

            // set pdf license
            var pdfLicense = new Aspose.Pdf.License();
            pdfLicense.SetLicense(totalLicense);

            // set word license
            var wordsLicense = new Aspose.Words.License();
            wordsLicense.SetLicense(totalLicense);

            // set excel license
            var cellsLicense = new Aspose.Cells.License();
            cellsLicense.SetLicense(totalLicense);

            // set powerpoint license
            var slidesLicense = new Aspose.Slides.License();
            slidesLicense.SetLicense(totalLicense);

            // sets Diagram(visio) license
            var visiolicense = new Aspose.Diagram.License();
            visiolicense.SetLicense(totalLicense);

            // sets Image license
            var imagelicense = new Aspose.Imaging.License();
            imagelicense.SetLicense(totalLicense);

            // sets OneNote License
            var noteLicense = new Aspose.Note.License();
            noteLicense.SetLicense(totalLicense);

            // sets email license
            var emailLicense = new Aspose.Email.License();
            emailLicense.SetLicense(totalLicense);
        }

        public int SectionOption(int fwkSectId)
        {
            int qryOptionCount;
            int MsgFile = default, childSec;
            string tocDocLabel, tocFileName;

            // =======================Gets count of options DOC in Section============================================
            qryOptionCount = _context.SUSPLAN_TOC_DOCUMENT.Count(t => t.SECTION_ID == qSectionId && t.DOC_OPTION == 1);
            // ===========End of Optional DOC count =====================

            // dtSectionDoc = Queries.qrySectionDocDT(fwkSectId, Application)	' Count of all docs in section
            var listTocDocument = _context.SUSPLAN_TOC_DOCUMENT.Where(t => t.SECTION_ID == qSectionId).ToList();

            // =================End of Doc Setion Query=======================

            // childSec = Queries.qryChildSection(fwkSectId, Application)
            var secHeader = _context.SUSPLAN_SEC_HDR.Find(qSectionId);
            if (secHeader != null)
                childSec = secHeader.CHILD_SECTION_ID ?? 0;
            else
                childSec = 0;

            // ===========End of Child Section Query ==================

            if (childSec != 0) // if there is a child section do not suppress the section.
            {
                return 0;
            }
            else if (qryOptionCount == listTocDocument.Count())    // all documents in section are optional
            {
                foreach (var item in listTocDocument)
                {
                    tocDocLabel = item.DOCUMENT_LABEL;

                    // =====Query to gets physical files name based on label and section id =========
                    var listDocument = _context.SUSPLAN_DOCUMENTS.Where(d => d.NODE_ID == svNodeId && d.DOCUMENT_LABEL.ToUpper() == tocDocLabel.ToUpper()).ToList();
                    tocFileName = "";

                    // ======End 
                    try
                    {
                        foreach (var obj in listDocument)
                            // assign the physical document name
                            tocFileName = obj.PHYSICAL_NAME;
                    }
                    catch // if Document label does not exists set variable it to null
                    {
                        tocFileName = "";
                    }

                    if (string.IsNullOrEmpty(tocFileName))  // if if physical file name is null increment counter by 1
                    {
                        MsgFile += 1;
                    }
                }
            }
            else
            {
                return 0;
            }

            if (MsgFile == qryOptionCount) // compares the missing file count to total optional documents
            {
                return 1; // suppress section
            }
            else
            {
                return 0;
            } // do not suppress
        }
        // add an entry to the rptPiecesTable

        public void AddToRPTable(ref string fwkTxtFileName, ref string fwkLabel, ref int fwkLevelNo, ref string fwkFilePath, ref string fwkSectionName)
        {
            rpct = rpct + 1; // add to the entry count
            rptPiecesTable[rpct, 1] = fwkTxtFileName; // store the file name
            rptPiecesTable[rpct, 2] = fwkLabel; // store the document label
            rptPiecesTable[rpct, 3] = fwkLevelNo.ToString(); // store the level number
            rptPiecesTable[rpct, 4] = fwkFilePath; // store the file path
            rptPiecesTable[rpct, 5] = fwkSectionName; // store the section name
        }

        // -----------------------------------------------------------------------------
        // This process is done every "long sleep" seconds.
        // Delete any entries that have been completed for completedPurgeDays days. 
        // -----------------------------------------------------------------------------
        public void CheckCompleted()
        {
            DateTime wkFinishedDate;
            long wkDaysFinished;
            LogMsg("Doing CheckCompleted", true);

            // Get all "completed" entries
            List<SUSPLAN_REPORT_Q> listReportQ = _context.SUSPLAN_REPORT_Q.Where(p => p.STATUS == "C").ToList();
            foreach (var item in listReportQ)
            {
                svRQId = item.RQ_ID;
                LogMsg("Processing Complete Report Queue Entry " + svRQId.ToString(), true);

                wkFinishedDate = item.FINISHED_DATE.HasValue ? item.FINISHED_DATE.Value : DateTime.Now;
                TimeSpan timeSpan = DateTime.Now - wkFinishedDate;
                wkDaysFinished = timeSpan.Days; // determine how many days it's been finished
                if (wkDaysFinished > myCaller.completedPurgeDays)                             // finished for more than 1 day?
                {
                    LogMsg("Deleting Complete Report Queue Entry " + svRQId.ToString() + " Complete For " + wkDaysFinished.ToString() + " Days. Instance: " + instName, false);
                    try
                    {
                        _context.SUSPLAN_REPORT_Q.Remove(item);
                        _context.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        LogMsg("*** Database Exception *** Cannot Delete Complete Report Queue Entry " + svRQId.ToString(), false, EventLogEntryType.Warning);
                        LogMsg("*** Message: " + ex.Message, false, EventLogEntryType.Warning);
                    }
                    DeleteTempFiles(svRQId);                     // delete any temp files
                }
            }

            LogMsg("Finished CheckCompleted", true);
        }

        // -----------------------------------------------------------------------------
        // This process is done every "long sleep" seconds.
        // Delete any entries that have been cancelled for cancelledPurgeDays days.  
        // -----------------------------------------------------------------------------
        public void CheckCancelled()
        {
            DateTime wkStartedDate;
            long wkDaysStarted;
            LogMsg("Doing CheckCancelled", true);

            // Get all "cancelled" entries

            List<SUSPLAN_REPORT_Q> listReportQ = _context.SUSPLAN_REPORT_Q.Where(p => p.STATUS == "X").ToList();
            foreach (var item in listReportQ)
            {
                svRQId = item.RQ_ID;
                LogMsg("Processing Cancelled Report Queue Entry " + svRQId.ToString(), true);
                wkStartedDate = item.STARTED_DATE.HasValue ? item.STARTED_DATE.Value : DateTime.Now;
                TimeSpan timeSpan = DateTime.Now - wkStartedDate;
                wkDaysStarted = timeSpan.Days; // determine how many days it started
                if (wkDaysStarted > myCaller.cancelledPurgeDays)                           // started for more than 1 day?
                {
                    LogMsg("Deleting Cancelled Report Queue Entry " + svRQId.ToString() + " In Queue For " + wkDaysStarted.ToString() + " Days. Instance: " + instName, false);
                    try
                    {
                        _context.SUSPLAN_REPORT_Q.Remove(item);
                        _context.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        LogMsg("*** Database Exception *** Cannot Delete Cancelled Report Queue Entry " + svRQId.ToString(), false, EventLogEntryType.Warning);
                        LogMsg("*** Message: " + ex.Message, false, EventLogEntryType.Warning);
                    }
                    DeleteTempFiles(svRQId);                     // delete any temp files
                }
            }
            LogMsg("Finished CheckCancelled", true);
        }

        // -----------------------------------------------------------------------------
        // This process is done every "long sleep" seconds.
        // Delete any entries that have been in Error status for errorPurgeDays days.
        // -----------------------------------------------------------------------------
        public void CheckErrors()
        {
            DateTime wkStartedDate;
            long wkDaysStarted;
            string wkXMLFileName;
            LogMsg("Doing CheckErrors", true);

            // Get all "error" entries
            List<SUSPLAN_REPORT_Q> listReportQ = _context.SUSPLAN_REPORT_Q.Where(p => p.STATUS == "E").ToList();
            foreach (var item in listReportQ)
            {
                svRQId = item.RQ_ID;
                LogMsg("Processing Error Report Queue Entry " + svRQId.ToString(), true);
                wkStartedDate = item.STARTED_DATE.HasValue ? item.STARTED_DATE.Value : DateTime.Now;
                TimeSpan timeSpan = DateTime.Now - wkStartedDate;
                wkDaysStarted = timeSpan.Days; // determine how many days it started
                if (wkDaysStarted > myCaller.errorPurgeDays)                           // started for more than 1 day?
                {
                    LogMsg("Deleting Error Report Queue Entry " + svRQId.ToString() + " In Queue For " + wkDaysStarted.ToString() + " Days. Instance: " + instName, false);
                    try
                    {
                        _context.SUSPLAN_REPORT_Q.Remove(item);
                        _context.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        LogMsg("*** Database Exception *** Cannot Delete Error Report Queue Entry " + svRQId.ToString(), false, EventLogEntryType.Warning);
                        LogMsg("*** Message: " + ex.Message, false, EventLogEntryType.Warning);
                    }
                    DeleteTempFiles(svRQId);                     // delete any temp files

                    // delete the error ticket on disk
                    wkXMLFileName = documentPath + @"\pdf_errors\" + GetJobTicketName() + "err";
                    if (File.Exists(wkXMLFileName))  // is there a job ticket?
                    {
                        File.Delete(wkXMLFileName);      // yep - delete it
                    }
                }
            }

            LogMsg("Finished CheckErrors", true);
        }

        // -----------------------------------------------------------------------------
        // This process is done every "long sleep" seconds.
        // Delete any entries that have been in "Pending", "Crystal Reports Generation",
        // or "Initializing" status for longPendingPurgeDays days.  
        // -----------------------------------------------------------------------------
        public void CheckLongPending()
        {
            DateTime wkStartedDate;
            long wkDaysStarted;
            string wkXMLFileName;
            LogMsg("Doing CheckLongPending", true);

            // Get all "long pending" entries
            List<SUSPLAN_REPORT_Q> listReportQ = _context.SUSPLAN_REPORT_Q.Where(p => p.STATUS == "P" || p.STATUS == "R" || p.STATUS == "I").ToList();
            foreach (var item in listReportQ)
            {
                svRQId = item.RQ_ID;
                LogMsg("Processing Report Queue Entry " + svRQId.ToString(), true);
                wkStartedDate = item.STARTED_DATE ?? DateTime.Now;
                TimeSpan timeSpan = DateTime.Now - wkStartedDate;
                wkDaysStarted = timeSpan.Days; // determine how many days it started
                if (wkDaysStarted > myCaller.longPendingPurgeDays)                           // started for more than 1 day?
                {
                    LogMsg("Deleting Long Pending/Crystal Report/Initializing Queue Entry " + svRQId.ToString() + " In Queue For " + wkDaysStarted.ToString() + " Days. Status = " + item.STATUS + " Instance: " + instName, false);

                    // delete the report q database record
                    try
                    {
                        _context.SUSPLAN_REPORT_Q.Remove(item);
                        _context.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        LogMsg("*** Database Exception *** Cannot Delete Long Pending/Crystal Report/Initializing Queue Entry " + svRQId.ToString(), false, EventLogEntryType.Warning);
                        LogMsg("*** Message: " + ex.Message, false, EventLogEntryType.Warning);
                    }
                    DeleteTempFiles(svRQId);                     // delete any temp files

                    // delete the job ticket on disk
                    wkXMLFileName = documentPath + @"\pdf_queue\" + GetJobTicketName();
                    if (File.Exists(wkXMLFileName))  // is there a job ticket?
                    {
                        File.Delete(wkXMLFileName);      // yep - delete it
                    }
                }
            }

            LogMsg("Finished CheckLongPending", true);
        }

        // -----------------------------------------------------------------------------
        // Delete all "temp" files for a given Report Queue ID
        // -----------------------------------------------------------------------------
        private void DeleteTempFiles(int fwkRQId)
        {
            string wkRQId9, wkRQIdStr, wkFileName;

            // build the (partial) file name for temp files to be deleted
            // format a 9 digit text report queue id
            wkRQIdStr = svRQId.ToString();                // get the report queue ID
            wkRQId9 = wkRQIdStr.PadLeft(9, '0');    // generate 9 character id
            wkFileName = "T" + wkRQId9;
            LogMsg("Deleting Files: " + wkFileName + "*.* in Directory " + documentPath + @"\temp", true);
            var objDi = new DirectoryInfo(documentPath + @"\temp");
            // Create an array representing the files in the directory.

            // check the files in the directory
            // Delete any for this job ticket
            var objFi = objDi.GetFiles();
            foreach (var fiTemp in objFi)
            {
                if (fiTemp.Name.StartsWith(wkFileName))
                {
                    LogMsg("Deleleting File: " + fiTemp.FullName, true);
                    try
                    {
                        fiTemp.Delete();
                    }
                    catch (Exception ex)
                    {
                        LogMsg("Exception Received From File Delete; Message: " + ex.Message, true);
                    }
                }
            }

            // wkFileName = "T" + wkRQId9 + "*.*"
            // wkFileName = documentPath + "\temp\" + wkFileName
            // LogMsg("Deleting Files: " + wkFileName, True)

            // ' delete any temp files
            // Try
            // File.Delete(wkFileName)
            // Catch ex As Exception    ' ignore file not found errors
            // LogMsg("Exception Received From File Delete; Message: " + ex.Message, True)
            // End Try
        }

        // ------------------------------------------------------------------------------
        // Build a job ticket name
        // ------------------------------------------------------------------------------
        private string GetJobTicketName()
        {
            string wkRQId9, wkRQIdStr, wkXMLFileName;
            wkRQIdStr = svRQId.ToString();                // get the report queue ID
            wkRQId9 = wkRQIdStr.PadLeft(9, '0');  // generate 9 character id
            wkXMLFileName = "P" + wkRQId9 + "_PDFCTL.xml";
            return wkXMLFileName;
        }

        // --------------------------------------------------
        // write messages to the log file and event log
        // --------------------------------------------------

        // informational message
        private void LogMsg(string msg, bool debugMsg)
        {
            DoLogMsg(msg, debugMsg, EventLogEntryType.Information);  // information message
        }

        // log a message using the passed log entry type
        public void LogMsg(string msg, bool debugMsg, EventLogEntryType logEntryType)

        {
            DoLogMsg(msg, debugMsg, logEntryType);
        }

        private void DoLogMsg(string msg, bool debugMsg, EventLogEntryType logEntryType)

        {
            string wkMsg;
            if (logFileInitialized)      // log file open?
            {
                if (logFileDay != (int)DateTime.Now.DayOfWeek)   // is it a new day?
                {
                    RolloverLogFile();                   // yep - close the current log file
                }                                  // and open a new one
            }

            if (debugMsg)
            {
                wkMsg = ">>> " + msg;
            }
            else
            {
                wkMsg = msg;
            }

            if (!debugMsg | myCaller.verbose)
            {
                try
                {
                    objEventLog.WriteEntry(wkMsg, logEntryType);
                }
                catch
                {
                    // no action - we don't want to hang up the service if there
                    // is an event log error
                }

                try
                {
                    wkMsg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + wkMsg;
                    wkMsg += Convert.ToString('\r');
                    if (logFileValid)            // got a valid log file?
                    {
                        objSWlog.WriteLine(wkMsg);   // yep - write message to log file
                    }
                    else
                    {
                        collLogFileQueue.Add(wkMsg);
                    } // nope - queue up messages
                }
                catch
                {
                    // no action - we don't want to hang up the service if there
                    // is a disk log error
                }
            }
        }

        // rollover processing for the log file at change of day
        // close the current log file
        // open the new log file
        private void RolloverLogFile()
        {
            if (!logFileValid)    // got a valid log file?
            {
                logFileDay = (int)DateTime.Now.DayOfWeek; // nope - save the new current day for rollover
                return;                // and exit
            }

            objSWlog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + ">>>>> Log Closed At End Of Day <<<<<" + '\r');
            objSWlog.Close();            // close the stream writer
            objFSlog.Close();            // close the file stream
            logFileDay = (int)DateTime.Now.DayOfWeek; // save the new current day for rollover
            logFileName = BuildLogFileName(logFileDay);        // get the log file name
            if (File.Exists(logFilePath + logFileName))   // does the file exist?
            {
                File.Delete(logFilePath + logFileName);       // yep - delete it
            }

            try
            {
                objFSlog = new FileStream(logFilePath + logFileName, FileMode.Create);
                objSWlog = new StreamWriter(objFSlog);
            }
            catch (Exception ex)
            {
                objEventLog.WriteEntry("Instance: " + instName + " Error Opening Log File. Log will not be written.  Message = " + ex.Message);
                logFileValid = false;    // don't try to use the log file
                return;
            }

            objSWlog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + ">>>>> Log Opened At Start Of Day <<<<<" + '\r');
        }

        public string Asposetest(string name, string loc, string ext)
        {
            string docName = Path.GetFileNameWithoutExtension(name); // remove extension from doc name
            string docOut = tempDir + docName + ".pdf";    // output document name
            //var SPinst = new SPCommon.SPCommon();

            // TODO: Replace generic error messages with more detailed ones.

            // convert extension to lowercase
            ext = ext.ToLower();

            // string arrays to hold valid doc types
            var spreadsheet = new string[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".ods", ".csv" };
            var wordDoc = new string[] { ".doc", ".dot", ".docm", ".dotm", ".docx", ".dotx", ".odt", ".rtf", ".txt" };
            var powerpoint = new string[] { ".ppt", ".pptx", ".pot", ".pps", ".potx", ".ppsx", ".odp", ".pptm" };
            var pdfs = new string[] { ".pdf", ".html", ".htm", ".pcl", ".epub", ".pdfa", ".svg" };
            var visio = new string[] { ".vsd", ".vsdx", ".vss", ".vst", ".vsx", ".vxt", ".vdw", ".vdx" };
            var crystal = new string[] { ".rpt" };
            var img = new string[] { ".bmp", ".jpeg", ".jpg", ".tiff", ".gif", ".png", ".psd", ".dxf", ".dwg" };
            var oneNote = new string[] { ".one" };
            var email = new string[] { ".msg", ".eml", ".emlx", ".mht" };
            var CvrError = new string[] { ".err" };
            if (spreadsheet.Contains(ext))  // checks if doc is spreadsheet
            {
                try
                {
                    var workbook = new Aspose.Cells.Workbook(loc + name);
                    workbook.Save(docOut, Aspose.Cells.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if crystal report failed
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing EXCEL File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorInd = "Y";
                    svErrorMsg = "** Error processing EXCEL File **";
                }    // error message passed to report Queue.
            }
            else if (wordDoc.Contains(ext))   // checks if doc is Word Doc
            {
                try
                {
                    var doc = new Aspose.Words.Document(loc + name);
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing WORD File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    svErrorInd = "Y";
                    svErrorMsg = "** Error processing WORD File **"; // error message passed to report Queue.
                    gotError = true;
                }
            }
            else if (powerpoint.Contains(ext)) // checks if doc is Power Point
            {
                try
                {
                    var slides = new Aspose.Slides.Presentation(loc + name);
                    slides.Save(docOut, Aspose.Slides.Export.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing SLIDES File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    svErrorInd = "Y";
                    svErrorMsg = "** Error processing SLIDES File **";   // error message passed to report Queue.
                    gotError = true;
                }
            }
            else if (pdfs.Contains(ext))   // checks if doc is PDF
            {
                try
                {
                    var docpdf = new Aspose.Pdf.Document();

                    // if html or htm must set load options that sets base path to SP root folder
                    if (ext == ".html" | ext == ".htm")
                    {
                        string basePath = documentPath;
                        var htmloptions = new Aspose.Pdf.HtmlLoadOptions(basePath);
                        docpdf = new Aspose.Pdf.Document(loc + name, htmloptions);
                    }
                    else
                    {
                        docpdf = new Aspose.Pdf.Document(loc + name);
                    }

                    docpdf.Save(docOut, Aspose.Pdf.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing PDF File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Error processing PDF File **";  // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (visio.Contains(ext))  // checks if doc is Visio
            {
                try
                {
                    var visioDoc = new Aspose.Diagram.Diagram(loc + name);
                    visioDoc.Save(docOut, Aspose.Diagram.SaveFileFormat.PDF);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing VISIO File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Error processing VISIO File **";    // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (crystal.Contains(ext))
            {
                if (string.IsNullOrEmpty(svSurveyId))
                {
                    svSurveyId = "0";
                }

                try
                {
                    ProcessCrystal(loc + name, docOut, svNodeId, Convert.ToInt32(svSurveyId), ref objEventLog);
                }
                catch (Exception ex)
                {
                    var doc = new Aspose.Words.Document();   // add blank pdf if crystal report failed
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing Crystal Report **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Crystal Report Error - " + ex.Message.ToString() + "**";  // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (img.Contains(ext))
            {
                Aspose.Pdf.Document pdf = new Aspose.Pdf.Document();//Add blank pdf
                var page = pdf.Pages.Add();// Add blank page

                // Retrieve names of all the Pdf files in a particular Directory
                string fileEntries = loc + name;

                // creat an image object
                try
                {
                    var image1 = new Aspose.Pdf.Image();
                    image1.File = fileEntries;
                    switch (ext ?? "")
                    {
                        case ".bmp":
                        case ".jpg":
                        case ".jpeg":
                        case ".tiff":
                        case ".tif":
                        case ".gif":
                        case ".png":
                            {
                                image1.FileType = Aspose.Pdf.ImageFileType.Base64;
                                break;
                            }

                        case ".psd":
                            {
                                // TODO: Fix
                                // image1.ImageInfo.ImageFileType = Aspose.Pdf.Generator.ImageFileType
                                throw new Exception("Unsupported");
                            }

                        case ".dxf":
                            {
                                // image1.ImageInfo.ImageFileType = Aspose.Pdf.Generator.ImageFileType.dx
                                throw new Exception("Unsupported");
                            }

                        case ".dwg":
                            {
                                // image1.ImageInfo.ImageFileType = Aspose.Pdf.Generator.ImageFileType.
                                throw new Exception("Unsupported");
                            }
                    }

                    image1.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;

                    // Create a BitMap object in order to get the information of image file
                    Bitmap myimage = new Bitmap(fileEntries);
                    // check if the width of the image file is greater than Page width or not
                    if (myimage.Width > page.PageInfo.Width)
                    {
                        // if the Image width is greater than page width, then set the page orientation to Landscape
                        page.PageInfo.IsLandscape = true;
                    }
                    else
                    {
                        // if the Image width is less than page width, then set the page orientation to Portrait
                        page.PageInfo.IsLandscape = false;
                    }
                    // add the image to paragraphs collection of the PDF document 
                    page.Paragraphs.Add(image1);

                    // save the PDF document
                    pdf.Save(docOut);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing IMAGE File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Error processing IMAGE File **";    // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (oneNote.Contains(ext))
            {
                try
                {
                    var Note = new Aspose.Note.Document(loc + name);
                    Note.Save(docOut, Aspose.Note.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing ONENOTE File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Error processing ONENOTE File **";  // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (email.Contains(ext))
            {
                try
                {
                    Aspose.Email.MailMessage emailAsp = Aspose.Email.MailMessage.Load(loc + name);
                    string tempMail = tempDir + docName + ".Mhtml";
                    Aspose.Email.SaveOptions emailOpt;
                    emailOpt = Aspose.Email.SaveOptions.DefaultMhtml;

                    // save email to Mhtml format
                    emailAsp.Save(tempMail, emailOpt);

                    // create an instance of LoadOptions and set the LoadFormat to Mhtml
                    var loadOptions = new Aspose.Words.LoadOptions();
                    loadOptions.LoadFormat = Aspose.Words.LoadFormat.Mhtml;

                    // load doc in Word with mhtl load options
                    var doc = new Aspose.Words.Document(tempMail, loadOptions);  // add blank pdf if Document failed to Save as PDF
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                }
                catch
                {
                    var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                    var builder = new Aspose.Words.DocumentBuilder(doc);

                    // Specify font formatting before adding text.
                    Aspose.Words.Font font = builder.Font;
                    font.Size = 16;
                    font.Bold = true;
                    font.Color = System.Drawing.Color.Black;
                    font.Name = "Verdana";
                    builder.Write("** Error processing Email File **");
                    doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);
                    gotError = true;
                    svErrorMsg = "** Error processing Email File **";    // error message passed to report Queue.
                    svErrorInd = "Y";
                }
            }
            else if (CvrError.Contains(ext))
            {
                var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                var builder = new Aspose.Words.DocumentBuilder(doc);

                // Specify font formatting before adding text.
                Aspose.Words.Font font = builder.Font;
                font.Size = 16;
                font.Bold = true;
                font.Color = System.Drawing.Color.Black;
                font.Name = "Verdana";
                builder.Write("Cover Page is missing from report folder");
                svErrorMsg = "** Cover Page is missing from report folder **";   // error message passed to report Queue.
                doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);

                // set error message to display in Document Queue.
                svErrorInd = "Y";
            }
            else     // invlaid doc type save to blank pdf doc
            {
                var doc = new Aspose.Words.Document(); // add blank pdf if Document failed to Save as PDF
                var builder = new Aspose.Words.DocumentBuilder(doc);

                // Specify font formatting before adding text.
                Aspose.Words.Font font = builder.Font;
                font.Size = 16;
                font.Bold = true;
                font.Color = Color.Black;
                font.Name = "Verdana";
                builder.Write(ext.ToString() + " is an invalid file type");
                svErrorMsg = "** " + ext.ToString() + " is an Invalid file type **";   // error message passed to report Queue.
                doc.Save(docOut, Aspose.Words.SaveFormat.Pdf);

                // set error message to display in Document Queue.
                svErrorInd = "Y";
            }

            return docOut;
        }

        // open the log file
        private bool OpenLogFile()
        {
            // Dim wkPath As String
            // Dim wkPos As Integer

            logFileDay = (int)DateTime.Now.DayOfWeek;    // save day for rollover

            // build the log file name and path
            // wkPath = configFileLoc.ToLower        ' get the configuration file directory
            // wkPos = wkPath.IndexOf("\susplanconfiguration")   ' look for our directory
            // If wkPos < 0 Then                           ' directory found?
            // wkPos = wkPath.IndexOf("\susplanadmin")   ' look for our directory
            // If wkPos < 0 Then                           ' directory found?
            // objEventLog.WriteEntry("Instance: " + instName + " Neither the SusplanConfiguration not the SusplanAdmin directory was found in configuration file path. Disk log file will not be written.")
            // Return False
            // End If

            // logFilePath = wkPath.Substring(0, wkPos) & "\documents\pdf_log\"  ' build the path
            // logFileName = BuildLogFileName(logFileDay)        ' get the log file name
            // Else
            // logFilePath = wkPath.Substring(0, wkPos) & "\SusplanDocuments\pdf_log\"  ' build the path
            // logFileName = BuildLogFileName(logFileDay)        ' get the log file name
            // End If

            logFileName = BuildLogFileName(logFileDay);       // get the log file name
            logFilePath = documentPath + @"\pdf_log\";  // build the path
            LogMsg("About to open log file: " + logFilePath + logFileName, true);
            try
            {
                objFSlog = new FileStream(logFilePath + logFileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                objSWlog = new StreamWriter(objFSlog);
            }
            catch (Exception ex)
            {
                objEventLog.WriteEntry("Instance: " + instName + " Error Opening Log File. Log will not be written. Message = " + ex.Message);
                return false;
            }

            objSWlog.WriteLine(">>>>>>>>>>>>>>>>> Log Opened <<<<<<<<<<<<<<<<<" + '\r');
            logFileValid = true;

            // write queued messages to the disk log
            logFileInitialized = true;
            foreach (string fwkMsg in collLogFileQueue)
                objSWlog.WriteLine(fwkMsg + '\r');
            collLogFileQueue = new List<string>();  // clear the queue
            return true;
        }

        // build the log file name
        private string BuildLogFileName(int thisDay)
        {
            string wkDayString = null;
            int wkDay;
            string wkFileName;
            wkDay = thisDay;
            switch (wkDay)                  // get the text day of the week
            {
                case (int)DayOfWeek.Sunday:
                    {
                        wkDayString = "Sunday";
                        break;
                    }

                case (int)DayOfWeek.Monday:
                    {
                        wkDayString = "Monday";
                        break;
                    }

                case (int)DayOfWeek.Tuesday:
                    {
                        wkDayString = "Tuesday";
                        break;
                    }

                case (int)DayOfWeek.Wednesday:
                    {
                        wkDayString = "Wednesday";
                        break;
                    }

                case (int)DayOfWeek.Thursday:
                    {
                        wkDayString = "Thursday";
                        break;
                    }

                case (int)DayOfWeek.Friday:
                    {
                        wkDayString = "Friday";
                        break;
                    }

                case (int)DayOfWeek.Saturday:
                    {
                        wkDayString = "Saturday";
                        break;
                    }
            }

            // build the log file name
            wkFileName = "Report_Monitor_Log_";
            wkFileName += wkDayString;
            wkFileName += ".log";
            return wkFileName;
        }

        // flush the log file to disk
        public void FlushLogFile()
        {
            if (logFileValid)
            {
                objSWlog.Flush();
            }
        }

        private void ProcessCrystal(string fwkRptFileName, string fwkOutputFileName, int folderID, int svSurveyid, [Optional, DefaultParameterValue(null)] ref EventLog EventWriter)
        {
            string fwkNodeDesc, fwkNodePath;
            CrystalDecisions.CrystalReports.Engine.ReportDocument objCrystalReport;

            // Use external event writer as it isn't set during init
            if (EventWriter != null)
            {
                objEventLog = EventWriter;
            }

            LogMsg("Processing Report: " + fwkRptFileName, false);
            LogMsg("Output Report: " + fwkOutputFileName, false);
            objCrystalReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            try
            {
                objCrystalReport.Load(fwkRptFileName, OpenReportMethod.OpenReportByTempCopy);  // load the report
            }
            catch (Exception ex)
            {
                LogMsg("Error encountered loading report " + objCrystalReport.FileName + " for entry " + svRQId.ToString() + " (First Try)", false, EventLogEntryType.Warning);
                LogMsg("Error Message: " + ex.Message + Convert.ToString('\r') + Convert.ToString('\n') + ex.StackTrace, false, EventLogEntryType.Warning);
            }

            if (!objCrystalReport.IsLoaded)
            {
                throw new Exception("Crystal Report Load Failed");
            }

            fwkNodeDesc = GetNodeDesc(folderID);     // get the node description
            fwkNodePath = GetNodePath(folderID);     // and the node path
            LogMsg("Node Id: " + svNodeId.ToString() + " Description: " + fwkNodeDesc + " Path: " + fwkNodePath, true);

            // Set the Report's parameter fields
            LogMsg(">>> Start Processing Report Parameters", true);
            DataDefinition objDataDefinition = objCrystalReport.DataDefinition;
            // get datadefinition obj for this report
            ParameterFieldDefinitions objParameterDefs = objDataDefinition.ParameterFields;
            ParameterValues myParameterCurrValues;
            // ***** Dim myParameterValue As ParameterDiscreteValue

            foreach (ParameterFieldDefinition myParameterDef in objParameterDefs)
            {
                object newValue;
                newValue = null;
                switch (myParameterDef.ParameterFieldName)
                {
                    case "svyid":
                        {
                            newValue = svSurveyid;
                            LogMsg("Parameter svyid: " + svSurveyid.ToString(), true);
                            break;
                        }

                    case "surveyid":
                        {
                            newValue = svSurveyid;
                            LogMsg("Parameter surveyid: " + svSurveyid.ToString(), true);
                            break;
                        }

                    case "nodeid":
                        {
                            newValue = folderID;
                            LogMsg("Parameter nodeid: " + folderID.ToString(), true);
                            break;
                        }

                    case "nodedesc":
                        {
                            newValue = fwkNodeDesc;
                            LogMsg("Parameter nodedesc: " + fwkNodeDesc, true);
                            break;
                        }

                    case "nodepath":
                        {
                            newValue = fwkNodePath;
                            LogMsg("Parameter nodepath: " + fwkNodePath, true);
                            break;
                        }

                    case "useextractdb":
                        {
                            LogMsg("Extract Database Will Be Used", true);
                            break;
                        }

                    default:
                        {
                            LogMsg("Unknown Parameter: " + myParameterDef.Name, true);
                            break;
                        }
                }

                if (newValue is object)
                {
                    // LogMsg("Storing Value For Parameter: " + .Name, True)
                    var myDiscreteParam = new ParameterDiscreteValue();
                    myDiscreteParam.Value = newValue;
                    myParameterCurrValues = myParameterDef.CurrentValues;
                    myParameterCurrValues.Add(myDiscreteParam);
                    myParameterDef.ApplyCurrentValues(myParameterCurrValues);
                }
            }

            LogMsg(">>> Finished Processing Report Parameters", true);

            // Set database connection properties
            LogMsg(">>> Start Processing Database Parameters", true);

            // get database connection parameters
            string fwkServerName = ConfigurationManager.AppSettings.Get("ServerName");
            string fwkDatabase = ConfigurationManager.AppSettings.Get("DBName");
            string fwkUserID = ConfigurationManager.AppSettings.Get("UserID");
            string fwkPassword = System.Web.HttpUtility.HtmlDecode(ConfigurationManager.AppSettings.Get("Password"));

            LogMsg("Applying Database Connection Information to Report", true);
            bool myRtn;
            myRtn = Logon(objCrystalReport, fwkServerName, fwkDatabase, fwkUserID, fwkPassword);
            LogMsg("Done Applying Database Connection Information to Report; Return Code: " + myRtn.ToString(), true);
            if (!myRtn)   // did logon fail?
            {
                throw new Exception("Could not change the report's database connection parameters.");
            }

            LogMsg(">>> Finished Processing Database Parameters", true);

            // export the report to disk
            LogMsg(">>> Starting Report Export To Disk", true);

            objCrystalReport.ExportToDisk(ExportFormatType.PortableDocFormat, fwkOutputFileName);

            LogMsg(">>> Finished Report Export To Disk", true);
            objCrystalReport.Close();
            objCrystalReport.Dispose();
            LogMsg("Finished Processing Report: " + fwkRptFileName, true);
        }

        private string GetNodeDesc(int fwkNodeId)
        {
            string rtnNodeDesc;
            var nodes = _context.SUSPLAN_NODES.Find(fwkNodeId);
            if (nodes != null)
            {
                rtnNodeDesc = nodes.DESCRIPTION;
            }
            else
            {
                rtnNodeDesc = "*Node Not Found*";
            }

            return rtnNodeDesc;
        }

        // ------------------------------------------------------------------------------
        // read up a node path to build the full node path
        // ------------------------------------------------------------------------------
        private string GetNodePath(int fwkNodeId)
        {
            string fwkNodePath;
            string fwkNodeDesc = null;
            int fwkParentNodeId;
            fwkNodePath = "";

            // read the subject node
            var nodes = _context.SUSPLAN_NODES.Find(fwkNodeId);
            if (nodes != null)
            {
                fwkParentNodeId = nodes.PARENT_ID ?? 0;
                if (fwkParentNodeId == 0)
                    fwkNodePath = "* Root Node *";
            }
            else
            {
                fwkNodePath = "*Node Not Found*";
                fwkParentNodeId = 0;
            }

            while (fwkParentNodeId != 0)              // read up the tree
            {
                var parentNodes = _context.SUSPLAN_NODES.Find(fwkParentNodeId);
                if (nodes != null)
                {
                    fwkParentNodeId = parentNodes.PARENT_ID ?? 0;
                    fwkNodeDesc = parentNodes.DESCRIPTION;
                }
                else
                {
                    fwkNodePath = "*Node In Chain Not Found*";
                    fwkParentNodeId = 0;
                }

                if (!string.IsNullOrEmpty(fwkNodePath))          // is there a lower node description?
                {
                    fwkNodePath = "." + fwkNodePath; // yes - add a separator
                }

                fwkNodePath = fwkNodeDesc + fwkNodePath; // add this node to the path
            }

            return fwkNodePath;           // return path to caller
        }

        private bool Logon(CrystalDecisions.CrystalReports.Engine.ReportDocument myReport, string strServer, string strDb, string strUserid, string strPass)
        {
            var myConnectionInfo = new CrystalDecisions.Shared.ConnectionInfo();
            myConnectionInfo.UserID = strUserid;
            myConnectionInfo.Password = strPass;
            myConnectionInfo.DatabaseName = strDb;
            // .Type = ConnectionInfoType.SQL
            myConnectionInfo.ServerName = strServer;

            //myReport.SetDatabaseLogon(strUserid, strPass, strServer, strDb);
            //myReport.VerifyDatabase()

            // apply database connection information to the main report
            if (!ApplyLogon(myReport, myConnectionInfo))
            {
                return false;

                // apply database connection information to any subreports
            }

            CrystalDecisions.CrystalReports.Engine.SubreportObject subobj;
            foreach (CrystalDecisions.CrystalReports.Engine.ReportObject obj in myReport.ReportDefinition.ReportObjects)
            {
                if (obj.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                {
                    subobj = (CrystalDecisions.CrystalReports.Engine.SubreportObject)obj;
                    if (!ApplyLogon(myReport.OpenSubreport(subobj.SubreportName), myConnectionInfo))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private bool ApplyLogon(CrystalDecisions.CrystalReports.Engine.ReportDocument myReport, CrystalDecisions.Shared.ConnectionInfo myConnectionInfo)

        {
            CrystalDecisions.Shared.TableLogOnInfo myLogonInfo;
            LogMsg("Applying Connection Info to Report: " + myReport.Name, true);
            LogMsg("Server Name: " + myConnectionInfo.ServerName, true);
            LogMsg("Database Name: " + myConnectionInfo.DatabaseName, true);
            LogMsg("UserID: " + myConnectionInfo.UserID, true);
            LogMsg("Password: " + myConnectionInfo.Password, true);
            foreach (CrystalDecisions.CrystalReports.Engine.Table myTable in myReport.Database.Tables)
            {
                LogMsg("Applying Connection Info to Table: " + myTable.Name, true);
                myLogonInfo = myTable.LogOnInfo;
                myLogonInfo.ConnectionInfo.ServerName = myConnectionInfo.ServerName;
                myLogonInfo.ConnectionInfo.DatabaseName = myConnectionInfo.DatabaseName;
                myLogonInfo.ConnectionInfo.UserID = myConnectionInfo.UserID;
                myLogonInfo.ConnectionInfo.Password = myConnectionInfo.Password;
                myTable.ApplyLogOnInfo(myLogonInfo);
                myTable.Location = myTable.Location; // could not change the database connect info without this ???!?"
                try
                {
                    myTable.TestConnectivity();
                }
                catch (Exception ex)
                {
                    LogMsg("Connection to database failed!", true);
                    LogMsg("Message: " + ex.Message, true);
                    return false;
                }
            }

            return true;
        }

    }
}