using System;
using System.IO;
using System.Net;
using System.Reflection;
using log4net;
using Newtonsoft.Json.Linq;

namespace Selenium_Automated_Testing.Utilities
{
    public class HtmlGenerator
    {
        private readonly string _strHtmlReportFile = "";
        //	Init the parameter for write log
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        //	Store value of color of last row in table report.
        private String _strLastRowColor = "#006600";
        public HtmlGenerator()
        {
        }
        public HtmlGenerator(string strReportFile)
        {
            _strHtmlReportFile = strReportFile;
        }

        /**********************************************************************************************************
	    '  Function Name:	initHTMLReport 
	    '  Purpose: 		This function will generate the report and some informatin in title of report. 
	    '  Inputs Parameters:
	    '  Returns: 
	    '   Creation Date: 11/09/2012
	    /**********************************************************************************************************/
        public void InitHtmlReport()
        {
            const string strApplication = "HarveyNash - Selenium Framework in C# - @@@";
            try
            {
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                var strDateTime = DateTime.Now.ToString();
                //Set Operation System value
                OperatingSystem os = Environment.OSVersion;
                var strOsVerionString = os.ToString();
                //To identify the system
                var strHostName = Dns.GetHostName();
                IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
                var strHostAddress = localIPs.GetLength(0) > 1 ? localIPs[1].ToString() : localIPs[0].ToString();

                StreamWriter writer = new StreamWriter(_strHtmlReportFile);
                writer.WriteLine("==================================================================================================================<BR>\n");
                writer.WriteLine("<font size=+1 color='#000000'>");
                writer.WriteLine("<B>Application:&nbsp</B>" + strApplication + "<BR>\n");
                writer.WriteLine("<B>OS		:&nbsp</B>" + strOsVerionString + "<BR>\n");
                writer.WriteLine("<B>Host Name :&nbsp</B>" + strHostName + "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp <B>Host IP:&nbsp</B>" + strHostAddress + "<BR>\n");
                writer.WriteLine("<B>Start time  :&nbsp</B>" + strDateTime + "<BR>\n");
                writer.WriteLine("</font>\n");
                writer.WriteLine("==================================================================================================================<BR><BR>\n");
                writer.WriteLine("<table cellspacing=0 border=1 width=100%>\n");
                writer.WriteLine("<tr>\n");
                writer.WriteLine("<td width='10%'><B>Test Case ID</B></td>\n");
                writer.WriteLine("<td width='50%'><B>Test Case Description</B></td>\n");
                writer.WriteLine("<td width='10%'><B>Status</B></td>\n");
                writer.WriteLine("<td width='40%'><B>Screenshot</B></td>\n");
                writer.WriteLine("<td width='20%'><B>Runtime</B></td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }

        /**********************************************************************************************************
	    '  Function Name:	createTableRowTestCase 
	    '  Purpose: 		This function will insert new row to next last row of table report as Test case. 
	    '  Inputs Parameters:
	    '  Returns: 
	    '   Creation Date: 11/09/2012
	    /**********************************************************************************************************/
        // ReSharper disable once InconsistentNaming
        public void CreateTableRowTestCase(String strTCID, String strTcDesc, bool bStatus, String strCapture)
        {
            var strRowColor = _strLastRowColor;
            try
            {
                StreamWriter writer = new StreamWriter(_strHtmlReportFile, true);
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                var strDateTime = DateTime.Now.ToString();
                writer.WriteLine("<tr>");
                writer.WriteLine("<td width=10% BGCOLOR=" + strRowColor + ">" + strTCID + "</td>\n");
                writer.WriteLine("<td width=50% BGCOLOR=" + strRowColor + ">" + strTcDesc + "</td>\n");
                string strStatus; //Value for Status PASS or FAIL
                string strColor; //Color for status column
                if (bStatus)
                {
                    strColor = "#000000";
                    strStatus = "Pass";
                }
                else
                {
                    strColor = "#FF0000";
                    strStatus = "Fail";
                }
                if (string.IsNullOrEmpty(strCapture))
                {
                    writer.WriteLine("<td width=10% BGCOLOR=" + strRowColor + "><font color=" + strColor + "><b>" + strStatus + "</b></font></td>\n");
                    strCapture = "&nbsp&nbsp";
                }

                else
                {
                    writer.WriteLine("<td width=10% BGCOLOR=" + strRowColor + "><font color=" + strColor + "><b>" + strStatus + "</b></font></td>\n");
                }
                writer.WriteLine("<td width=40% BGCOLOR=" + strRowColor + ">" + strCapture + "</td>\n");
                writer.WriteLine("<td width='20%'BGCOLOR=" + strRowColor + ">" + strDateTime + "</td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
                _strLastRowColor = strRowColor.Equals("#C8B560") ? "#ECD672" : "#C8B560";
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }

        /**********************************************************************************************************
	    '  Function Name:	createTableRowTestCase 
	    '  Purpose: 		This function will insert new row to next last row of table report as Test suite. 
	    '  Inputs Parameters:
	    '  Returns: 
	    '   Creation Date: 11/09/2012
	    /**********************************************************************************************************/
        // ReSharper disable once InconsistentNaming
        // ReSharper disable once InconsistentNaming
        public void CreateTableRowTestSuite(String strTSID, String strTsDesc)
        {
            const string strRowColor = "#C8D000";
            const string strColor = "#FF0000";
            try
            {
                StreamWriter writer = new StreamWriter(_strHtmlReportFile, true);
                writer.WriteLine("<tr>");
                writer.WriteLine("<td width='10%' BGCOLOR=" + strRowColor + "><font color='" + strColor + "'><B>" + strTSID + "</B></td>\n");
                writer.WriteLine("<td colspan='3' width='90%' BGCOLOR=" + strRowColor + "><font color='" + strColor + "'><B>" + strTsDesc + "</B></td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }

        /**********************************************************************************************************
        '  Function Name:	createTableRowResultTestSuite 
        '  Purpose: 		This function will insert new row to next last row of table report as Test suite result. 
        '  Inputs Parameters: JSONObject joResultTestSuite -- This store the result of testsuite as structure below.
                {
                    "pass": <int>,
                    "fail": "<int>,
                    "notExecute": <int>,
                    "totalTestcase": <int>
                }
        '  Returns: 
        '   Creation Date: 11/09/2012
        /**********************************************************************************************************/
        public void CreateTableRowResultTestSuite(JObject joResultTestSuite)
        {
            const string strRowColor = "#339966";
            const string strColor = "#000000";
            try
            {
                //Get the result of Test suite from JSONObject
                int intNumberOfPass = (int)joResultTestSuite["pass"];
                int intNumberOfFail = (int)joResultTestSuite["fail"];
                int intNumberOfNotExecute = (int)joResultTestSuite["notExecute"];

                var strDataPrint = "Pass: " + intNumberOfPass + " --- Fail: " + intNumberOfFail +
                                      " --- Not Execute: " + intNumberOfNotExecute;

                StreamWriter writer = new StreamWriter(_strHtmlReportFile, true);
                writer.WriteLine("<tr>");
                writer.WriteLine("<td colspan='4' width='100%' align='right' BGCOLOR=" + strRowColor + "><font color='" + strColor + "'><B>" + strDataPrint + "</B></td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }

        /**********************************************************************************************************
	    '  Function Name:	createTableRowFinalResult 
	    '  Purpose: 		This function will insert new row to next last row of table report as Test suite result. 
	    '  Inputs Parameters: JSONObject joResultTestSuite -- This store the result of testsuite as structure below.
			    {
				    "pass": <int>,
				    "fail": "<int>,
				    "notExecute": "<int>
			    }
	    '  Returns: 
	    '   Creation Date: 11/09/2012
	    /**********************************************************************************************************/
        public void CreateTableRowFinalResult(JObject joResultAllTestSuites)
        {
            const string strRowColor = "#FFFFFF";
            const string strColor = "#000000";
            try
            {
                //Get the result of Test suite from JSONObject
                int intNumberOfPass = (int)joResultAllTestSuites["PASS"];
                int intNumberOfFail = (int)joResultAllTestSuites["FAIL"];
                int intNumberOfNotExecute = (int)joResultAllTestSuites["notExecute"];
                int intNumberOfTotalTestCase = intNumberOfPass + intNumberOfFail + intNumberOfNotExecute;

                var strDataPrint = "Total Testcase: " + intNumberOfTotalTestCase + " ----- Pass: " +
                                      intNumberOfPass + " --- Fail: " + intNumberOfFail +
                                      "--- Not Execute: " + intNumberOfNotExecute;

                StreamWriter writer = new StreamWriter(_strHtmlReportFile, true);
                writer.WriteLine("<tr> &nbsp</tr>\n");
                writer.WriteLine("<tr>");
                writer.WriteLine("<td colspan='4' width='100%' align='right' BGCOLOR=" + strRowColor + "><font color='" + strColor + "'><B>" + strDataPrint + "</B></td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
        }

        /**********************************************************************************************************
        '  Function Name:	insertEndTime 
        '  Purpose: 		This function will insert the End Time of test at the end of report 
        '  Inputs Parameters: JSONObject joResultTestSuite -- This store the result of testsuite as structure below.
                {
                    "pass": <int>,
                    "fail": "<int>,
                    "notExecute": "<int>
                }
        '  Returns: 
        '   Creation Date: 11/09/2012
        /**********************************************************************************************************/
        public void InsertEndTime()
        {
            const string strRowColor = "#FFFFFF";
            const string strColor = "#000000";
            try
            {
                // ReSharper disable once SpecifyACultureInStringConversionExplicitly
                var strDateTime = DateTime.Now.ToString();
                StreamWriter writer = new StreamWriter(_strHtmlReportFile, true);
                writer.WriteLine("<tr> &nbsp</tr>\n");
                writer.WriteLine("<tr>");
                writer.WriteLine("<td colspan='4' width='100%' align='right' BGCOLOR=" + strRowColor + "><font color='" + strColor + "'>End Time: " + strDateTime + "</td>\n");
                writer.WriteLine("</tr>\n");
                writer.Close();
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }
    }
}
