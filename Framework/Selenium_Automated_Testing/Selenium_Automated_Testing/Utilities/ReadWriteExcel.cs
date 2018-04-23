using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
namespace Selenium_Automated_Testing.Utilities
{
    public class ReadWriteExcel
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string strTestCaseFile = "";
        string strDataFile = "";
        string strResultFile = "";
        private Common common = new Common();
        public ReadWriteExcel()
        {
            strTestCaseFile = common.StrTestCaseFile;
            strDataFile = common.StrDataFile;
            strResultFile = common.StrResultFile;
            CopyFile();
        }
        //Purpose: get filename from full path
        public string GetFileName(string strFileName)
        {
            string strOnlyFileName = "";
            try
            {
                char[] delimiters = new char[] { '\\' };
                string[] arrFileName = strFileName.Split(delimiters);
                // Get the final value in array that is the File Name.
                int intMax = arrFileName.GetLength(0);
                strOnlyFileName = arrFileName[intMax - 1];
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return strOnlyFileName;
        }
        //Purpose: This function will Copy TestCase File to Write the Result to it
        private void CopyFile()
        {
            bool bOverWritten = true;
            try
            {
                File.Copy(strTestCaseFile, strResultFile, bOverWritten);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
        }
        //Purpose: This function will read all rows of specific sheet with limit of column
        public string[,] ReadInputData(string strFileName, string strSheetName, int intColumn)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            string[,] arrData = null;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strSheetName];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int intRowCount = xlRange.Rows.Count;
                int intColumnCount = xlRange.Columns.Count;
                arrData = new string[intRowCount - 1, intColumn];
                // PROCESS EACH SPREADSHHET ROW
                int intR = 0;
                int intC = 0;
                for (int i = 2; i <= intRowCount; i++, intR++)
                {
                    // PROCESS EACH COLUMN IN EACH ROW
                    for (int j = 1; j <= intColumn; j++, intC++)
                    {
                        string strTemp = Convert.ToString(xlRange.Cells[i, j].Value2);    //  HAD TO DO THIS BECAUSE OF DYNAMIC BINDING TO GET
                        if (!string.IsNullOrEmpty(strTemp))
                            strTemp = strTemp.Trim();
                        //  STRING OF THE   VALUE OF THE CELL
                        arrData[intR, intC] = strTemp;
                    }
                    intC = 0;
                }

                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return arrData;
        }
        //Purpose: This function will read all rows of specific sheet with all column
        public string[,] ReadInputData(string strFileName, string strSheetName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            string[,] arrData = null;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strSheetName];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int intRowCount = xlRange.Rows.Count;
                int intColumnCount = xlRange.Columns.Count;
                arrData = new string[intRowCount, intColumnCount];

                // PROCESS EACH SPREADSHHET ROW
                int intR = 0;
                int intC = 0;
                for (int i = 1; i <= intRowCount; i++, intR++)
                {
                    // PROCESS EACH COLUMN IN EACH ROW
                    for (int j = 1; j <= intColumnCount; j++, intC++)
                    {
                        string strTemp = Convert.ToString(xlRange.Cells[i, j].Value2);    //  HAD TO DO THIS BECAUSE OF DYNAMIC BINDING TO GET
                        if (!string.IsNullOrEmpty(strTemp))
                            strTemp = strTemp.Trim();
                        //  STRING OF THE   VALUE OF THE CELL
                        arrData[intR, intC] = strTemp;
                    }
                    intC = 0;
                }
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return arrData;
        }
        //Purpose: This function will get the position of COLUMN base on specific value and sheet name
        public int getPositionOfColumn(string strFileName, string strSheetName, string strValue)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            int intPositionOfColumn = 99999;
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strSheetName];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                int intColumn = xlWorksheet.UsedRange.Columns.Count;  // Store the number of column in excel sheet

                for (int i = 1; i <= intRow; i++)
                {
                    for (int j = 1; j <= intColumn; j++)
                    {
                        string strActualValue = Convert.ToString(xlRange.Cells[i, j].Value2);    //  HAD TO DO THIS BECAUSE OF DYNAMIC BINDING TO GET
                        if (!string.IsNullOrEmpty(strActualValue))
                        {
                            strActualValue = strActualValue.Trim();
                            if (strActualValue.Equals(strValue))
                            {
                                intPositionOfColumn = j;
                                break;
                            }
                        }
                    }
                    if (intPositionOfColumn != 99999)
                    {
                        break;
                    }
                }
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return intPositionOfColumn;
        }
        //Purpose: This function will get the position of ROW base on specific value and sheet name
        public int getPositionOfRow(string strFileName, string strSheetName, string strValue)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            int intPositionOfRow = 99999;
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strSheetName];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                int intColumn = xlWorksheet.UsedRange.Columns.Count;  // Store the number of column in excel sheet
                for (int i = 1; i <= intColumn; i++)
                {
                    for (int j = 1; j <= intRow; j++)
                    {
                        string strActualValue = Convert.ToString(xlRange.Cells[j, i].Value2);    //  HAD TO DO THIS BECAUSE OF DYNAMIC BINDING TO GET
                        if (!string.IsNullOrEmpty(strActualValue))
                        {
                            strActualValue = strActualValue.Trim();

                            if (strActualValue.Equals(strValue))
                            {
                                intPositionOfRow = j;
                                break;
                            }
                        }
                    }
                    if (intPositionOfRow != 99999)
                    {
                        break;
                    }
                }
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return intPositionOfRow;
        }
        //Purpose: This function will get all TestSuites that Execute = Y then return.
        public JArray GetTestSuitesToExecute()
        {
            int intColumn = 5; //(0)Execute		(1)TestSuiteID		(2)Description	(3)SuiteSheet	(4)TestCaseSheet            
            string[,] arrData = ReadInputData(strTestCaseFile, common.StrSuiteSheet, intColumn);
            JArray jaTestSuitesExecute = new JArray();  // This to store list of Test Suites will be execute.          
            try
            {
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseID = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strSuiteSheet = arrData[intR, 3];
                    string strTestCaseSheet = arrData[intR, 4];

                    if (strExecute.Equals("Y"))
                    {
                        JObject joTestSuite = new JObject();
                        joTestSuite.Add("execute", strExecute);
                        joTestSuite.Add("testsuiteID", strTestCaseID);
                        joTestSuite.Add("description", strDescription);
                        joTestSuite.Add("suiteSheet", strSuiteSheet);
                        joTestSuite.Add("testcaseSheet", strTestCaseSheet);
                        jaTestSuitesExecute.Add(joTestSuite);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return jaTestSuitesExecute;
        }
        //Purpose: This function will get all TestSuites in TestSuite sheet then return the list of TestSuites
        public JArray GetListOfTestSuites()
        {
            int intColumn = 2; //(0)Execute		(1)TestSuiteID
            string[,] arrData = ReadInputData(strTestCaseFile, common.StrSuiteSheet, intColumn);

            JArray jaTestSuitesExecute = new JArray();  // This to store list of Test Suites will be execute.          
            
            try
            {
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestSuiteID = arrData[intR, 1];
                    if (!string.IsNullOrEmpty(strTestSuiteID))
                    {
                        JObject joTestSuite = new JObject();
                        joTestSuite.Add("execute", strExecute);
                        joTestSuite.Add("testsuiteID", strTestSuiteID);
                        jaTestSuitesExecute.Add(joTestSuite);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return jaTestSuitesExecute;
        }
        //Purpose: This function will get all TestCases that Execute = Y then return it in JArray
        public JArray GetTestCasesToExecute(string strSuiteSheet)
        {
            JArray jaTestCasesExecute = new JArray();	//Store the list of test cases will be execute
            int intColumn = 4; //(0)Execute		(1)TestCase		(2)Description	(3)DataID
            try
            {
                string[,] arrData = ReadInputData(strTestCaseFile, strSuiteSheet, intColumn);
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseID = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strDataID = arrData[intR, 3];
                    if (strExecute.Equals("Y"))
                    {
                        JObject joInputData = new JObject();
                        joInputData.Add("execute", strExecute);
                        joInputData.Add("testcaseID", strTestCaseID);
                        joInputData.Add("description", strDescription);
                        joInputData.Add("dataID", strDataID);
                        jaTestCasesExecute.Add(joInputData);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return jaTestCasesExecute;
        }
        //Purpose: This function will get all TestCases in specific TestSuite sheet then return the list of TestCases
        public JArray GetListOfTestCases(string strSuiteSheet)
        {
            JArray jaTestCasesExecute = new JArray();	//Store the list of test cases will be execute
            /*	jaTestCaseExecute
            [{
                  "execute": "<string>",
                  "testcaseID": "<string>",            
            }, ..., ]
            */
            int intColumn = 2; //(0)Execute		(1)TestCase
            try
            {
                string[,] arrData = ReadInputData(strTestCaseFile, strSuiteSheet, intColumn);

                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseID = arrData[intR, 1];
                    if (!string.IsNullOrEmpty(strTestCaseID))
                    {
                        JObject joInputData = new JObject();
                        joInputData.Add("execute", strExecute);
                        joInputData.Add("testcaseID", strTestCaseID);
                        jaTestCasesExecute.Add(joInputData);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }

            return jaTestCasesExecute;
        }
        //Purpose: This function will get all TestCases that Execute = Y then return it in a string with only TestCaseID in it.
        public string GetListOfTestCasesToExecute(string strSuiteSheet)
        {
            string strTestCaseExecute = "";
            JArray jaExecuteTestcases = new JArray();	//["testcaseID 001", "testcaseID 002" "testcaseID 003", ....]
            int intColumn = 2; //(0)Execute		(1)TestCase
            try
            {
                string[,] arrData = ReadInputData(strTestCaseFile, strSuiteSheet, intColumn);

                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseID = arrData[intR, 1];

                    if (strExecute.Equals("Y"))
                    {
                        jaExecuteTestcases.Add(strTestCaseID);
                    }
                }
                strTestCaseExecute = jaExecuteTestcases.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return strTestCaseExecute;
        }

        public JArray GetTestSteps(string strSheetName, string strTestCaseExecute)
        {
            JArray jaExecuteSteps = new JArray();
            /*
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>",
                "description": "<string>",
                "keyword": "<string>",
                "object": "<string>",
                "dataObject": "<string>",
            }, ..., ]
            */
            int intColumn = 6; //(0)TestCaseID	(1)StepID	(2)Description	(3)Keyword	(4)Object	(5)DataObject
            try
            {
                string[,] arrData = ReadInputData(strTestCaseFile, strSheetName, intColumn);
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strTestCaseID = arrData[intR, 0];
                    string strTestStepID = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strKeyword = arrData[intR, 3];
                    string strObject = arrData[intR, 4];
                    string strDataObject = arrData[intR, 5];

                    if (strTestCaseExecute.Equals(strTestCaseID))
                    {
                        JObject joTestcase = new JObject();
                        joTestcase.Add("testcaseID", strTestCaseID);
                        joTestcase.Add("teststepID", strTestStepID);
                        joTestcase.Add("description", strDescription);
                        joTestcase.Add("keyword", strKeyword);
                        joTestcase.Add("object", strObject);
                        joTestcase.Add("dataObject", strDataObject);

                        jaExecuteSteps.Add(joTestcase);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return jaExecuteSteps;
        }
        public JArray GetListOfTestSteps(String strTestCaseSheet)
        {
            JArray jaTestSteps = new JArray();
            /*
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>"                
            }, ..., ]
            */
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strTestCaseFile, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strTestCaseSheet];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int intRowCount = xlRange.Rows.Count;

                // PROCESS EACH SPREADSHHET ROW                
                for (int i = 2; i <= intRowCount; i++)
                {
                    /*	(0)TestCaseID	(1)StepID	(2)Description
				 	    Here we just get TestcaseID and StepID
				    */
                    string strTCID = Convert.ToString(xlRange.Cells[i, 1].Value2);
                    string strTSID = Convert.ToString(xlRange.Cells[i, 2].Value2);
                    JObject joTestStep = new JObject();
                    if (!string.IsNullOrEmpty(strTCID))
                    {
                        joTestStep.Add("testcaseID", strTCID);
                        joTestStep.Add("teststepID", strTSID);
                        jaTestSteps.Add(joTestStep);
                    }
                }
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return jaTestSteps;
        }
        public JArray GetListOfTestStepsExecuted(String strSuiteSheet, String strTestCaseSheet)
        {
            JArray jaTestSteps = new JArray();
            /*
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>"                
            }, ..., ]
            */
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            JArray jaTestCaseExecute = GetTestCasesToExecute(strSuiteSheet);
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strTestCaseFile, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strTestCaseSheet];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int intRowCount = xlRange.Rows.Count;
                int intColumnCount = xlRange.Columns.Count;
                // PROCESS EACH SPREADSHHET ROW                
                for (int i = 2; i <= intRowCount; i++)
                {
                    /*	(0)TestCaseID	(1)StepID	(2)Description
				 	    Here we just get TestcaseID and StepID
				    */
                    string strTestCaseID = Convert.ToString(xlRange.Cells[i, 1].Value2);
                    // PROCESS EACH COLUMN IN EACH ROW
                    for (int j = 1; j <= jaTestCaseExecute.Count; j++)
                    {
                        JObject joTestCaseExecuted = (JObject)jaTestCaseExecute[j - 1];
                        string strTestCaseExecuted = (string)(joTestCaseExecuted["testcaseID"]);

                        if (strTestCaseID.Equals(strTestCaseExecuted))
                        {
                            string strTCID = Convert.ToString(xlRange.Cells[i, 1].Value2);
                            string strTSID = Convert.ToString(xlRange.Cells[i, 2].Value2);
                            JObject joTestStep = new JObject();
                            joTestStep.Add("testcaseID", strTCID);
                            joTestStep.Add("teststepID", strTSID);
                            jaTestSteps.Add(joTestStep);
                            break;	//Break for loop
                        }
                    }
                }
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return jaTestSteps;
        }

        public int GetNumberOfColumn(string strFileName, string strSheetName)
        {
            int intColumn = 0;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(strSheetName);
                intColumn = xlWorkSheet.UsedRange.Columns.Count;
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return intColumn;
        }
        public int GetNumberOfRow(string strFileName, string strSheetName)
        {
            int intRow = 0;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = true;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(strSheetName);
                intRow = xlWorkSheet.UsedRange.Rows.Count;
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
            return intRow;
        }
        public JObject GetDataForEachTestCase(String strSheetName, String strDataID, JArray jaTestSteps, String[] arrDataObjects)
        {
            JObject joExecuteData = new JObject();
            /*
            {
                "dataObjectName" : "value",...	  	  			
            }
            */
            try
            {
                string[,] arrData = ReadInputData(strDataFile, strSheetName);

                int intLength = arrDataObjects.Length;
                int intNumberOfColumnData = GetNumberOfColumn(strDataFile, strSheetName);

                int intIndexOfRow = 0; //Store the index of row data
                //Get index of ROW to get data
                for (int i = 0; i < arrData.GetLength(0); i++)
                {
                    if (strDataID.Equals(arrData[i, 0]))
                    {
                        intIndexOfRow = i;
                        break;
                    }
                }
                //Get Data and input to JSON Object
                for (int i = 0; i < intLength; i++)
                {
                    string strDataObjectName = arrDataObjects[i];
                    for (int j = 0; j < intNumberOfColumnData; j++)
                    {
                        string strHeader = arrData[0, j].ToString();
                        strHeader = strHeader.Trim();
                        if (strHeader.Equals(strDataObjectName))
                        {
                            if (string.IsNullOrEmpty((string)joExecuteData[strDataObjectName]))
                            {
                                string strData = arrData[intIndexOfRow, j];
                                joExecuteData.Add(strDataObjectName, strData);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            return joExecuteData;
        }
        public void WriteResultTestCaseDetails(String strSuiteSheet, String strTestCaseSheet, JObject joResultTestCaseDetails)
        {
            /*	joResultTestSuiteDetails
            {
                "tcID":	[
                            {
                                "tsID": <string>,
                                "result": "<boolean>"
                            }, ......,	
                        ], ...., 
            }
            */
            JArray jaTestStepsExecuted = new JArray();	// This will store the list TC that execute = Y
            /*
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>"
            }, ...., ]
            */
            JArray jaTestSteps = new JArray();	// This will store the list all TCs in TEST CASE DETAILS sheet
            /*
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>"
            }, ...., ]
            */
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = false;
            xlApp = new Excel.Application();
            int intColumnResult = common.IntColumnOfResultTestCaseDetails;	// get the position of Column result in test case details sheet of test case file
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strResultFile, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strTestCaseSheet];
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                jaTestSteps = GetListOfTestSteps(strTestCaseSheet);
                jaTestStepsExecuted = GetListOfTestStepsExecuted(strSuiteSheet, strTestCaseSheet);
                string strTestCaseExecute = GetListOfTestCasesToExecute(strSuiteSheet);
                //Go through all row of test details.
                for (int k = 0; k < intRow - 1; k++)
                {
                    JObject joTestCase = (JObject)jaTestSteps[k];
                    string strTCID = (string)joTestCase["testcaseID"];
                    string strTSID = (string)joTestCase["teststepID"];
                    if (strTestCaseExecute.Contains(strTCID))
                    {
                        int intLenght = jaTestStepsExecuted.Count;
                        //Go through all row of test details that had been executed. 
                        for (int i = 0; i < intLenght; i++)
                        {
                            JObject joTestCaseExecuted = (JObject)jaTestStepsExecuted[i];
                            string strTCIDExecuted = (string)joTestCaseExecuted["testcaseID"];
                            string strTSIDExecuted = (string)joTestCaseExecuted["teststepID"];
                            if (strTCID.Equals(strTCIDExecuted) && strTSID.Equals(strTSIDExecuted))
                            {
                                JArray jaTestStepsResult = new JArray();
                                jaTestStepsResult = (JArray)joResultTestCaseDetails[strTCIDExecuted];
                                /*[
                                    {
                                        "tsID": <string>,
                                        "result": "<boolean>"
                                    }, ......,	
                                ]*/
                                int intLenghtOfTC = jaTestStepsResult.Count;
                                for (int j = 0; j < intLenghtOfTC; j++)
                                {
                                    JObject joTestStepResult = (JObject)jaTestStepsResult[j];
                                    string strTSIDActual = (string)joTestStepResult["tsID"];
                                    bool strResult = (bool)joTestStepResult["result"];
                                    string strResultToWrite = "";

                                    if (strTSIDExecuted.Equals(strTSIDActual))	// Test step match with result
                                    {
                                        //Set Format for FONT                                                                                
                                        xlWorksheet.Cells[k + 2, intColumnResult].Interior.Color = System.Drawing.Color.White;
                                        xlWorksheet.Cells[k + 2, intColumnResult].Font.Bold = true;
                                        if (strResult)
                                        {
                                            xlWorksheet.Cells[k + 2, intColumnResult].Font.Color = System.Drawing.Color.Blue;
                                            strResultToWrite = "PASS";
                                        }
                                        else
                                        {
                                            xlWorksheet.Cells[k + 2, intColumnResult].Font.Color = System.Drawing.Color.Red;
                                            strResultToWrite = "FAIL";
                                        }
                                        //Set Border for CELL                                        
                                        //Set value for CELL
                                        xlWorksheet.Cells[k + 2, intColumnResult] = strResultToWrite;
                                        //Console.WriteLine("Write value at {0} and {1}", k + 2, intColumnResult);
                                        break;
                                    }	// End if
                                }	// Exit For
                                break;
                            }	// End if
                        }	//	Exit for walkthrought all test case that executed	
                    }	// End if
                }	// Exit for walkthrought all row in Test case details sheet.
                xlWorkBook.RefreshAll();
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
        }

        public void WriteResultTestSuiteDetails(String strSuiteSheet, JObject joResultTestSuiteDetails)
        {
            /*	joResultTestSuiteDetails
                {
                    "tcID" <string>: "result" <boolean>, ....
                }
            */
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = false;
            xlApp = new Excel.Application();
            JArray jaTestCases = new JArray();	// This will store the list all TCs in TEST SUITE DETAILS sheet
            /*
            [	{
		 
                    "execute": "<string>",
                    "testcaseID": "<string>"`
                }
            , ...., ]
            */
            int intColumnResult = common.IntColumnOfResultTestSuiteDetails;	// get the position of Column result in test suite details sheet of test case file            	

            try
            {

                xlWorkBook = xlApp.Workbooks.Open(strResultFile, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strSuiteSheet];
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                jaTestCases = GetListOfTestCases(strSuiteSheet);
                //Go through all row of test suite details.
                for (int k = 0; k < intRow - 1; k++)
                {
                    JObject joTestCase = (JObject)jaTestCases[k];
                    string strExecute = (string)joTestCase["execute"];
                    string strTCID = (string)joTestCase["testcaseID"];
                    if (strExecute.Equals("Y"))
                    {
                        string strResultToWrite;
                        bool strResult = (bool)joResultTestSuiteDetails[strTCID];
                        //Set Format for FONT                                                                                
                        xlWorksheet.Cells[k + 2, intColumnResult].Interior.Color = System.Drawing.Color.White;
                        xlWorksheet.Cells[k + 2, intColumnResult].Font.Bold = true;
                        if (strResult)
                        {
                            xlWorksheet.Cells[k + 2, intColumnResult].Font.Color = System.Drawing.Color.Blue;
                            strResultToWrite = "PASS";
                        }
                        else
                        {
                            xlWorksheet.Cells[k + 2, intColumnResult].Font.Color = System.Drawing.Color.Red;
                            strResultToWrite = "FAIL";
                        }
                        //Set value for CELL
                        xlWorksheet.Cells[k + 2, intColumnResult] = strResultToWrite;
                        //Console.WriteLine("Write value at {0} and {1}", k + 2, intColumnResult);
                    }
                }	// Exit for walkthrought all row in Test case details sheet.
                xlWorkBook.RefreshAll();
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
        }
        public void WriteResultTestSuitesSheet(String strTestSuitesSheet, JObject joResultAllTestSuites)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = false;
            xlApp = new Excel.Application();
            JArray jaTestSuites = new JArray();	// This will store the list all test suites in TestSuites sheet
            /*
            [	{
                    "execute": "<string>",
                    "testsuiteID": "<string>"`
                }
            , ...., ]
            */
            int intColumnResult = common.IntColumnOfPass;	// get the position of Column Pass in testsuites sheet of test case file		    
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strResultFile, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[strTestSuitesSheet];
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                jaTestSuites = GetListOfTestSuites();
                //Go through all row of testsuites sheet
                for (int k = 0; k < intRow - 1; k++)
                {
                    JObject joTestCase = (JObject)jaTestSuites[k];
                    string strExecute = (string)joTestCase["execute"];
                    string strTestSuiteID = (string)joTestCase["testsuiteID"];
                    if (strExecute.Equals("Y"))
                    {
                        JObject joTestSuiteResult = new JObject();
                        //{
                        //    "pass": <int>,
                        //    "fail": <int>,
                        //    "notExecute": <int>
                        //}
                        joTestSuiteResult = (JObject)joResultAllTestSuites[strTestSuiteID];
                        int intNumberOfPass = (int)joTestSuiteResult["PASS"];
                        int intNumberOfFail = (int)joTestSuiteResult["FAIL"];
                        int intNumberOfNotExecute = (int)joTestSuiteResult["notExecute"];
                        //Set Format for FONT
                        for (int i = 0; i < 3; i++)
                        {
                            xlWorksheet.Cells[k + 2, intColumnResult + i].Interior.Color = System.Drawing.Color.White;
                            xlWorksheet.Cells[k + 2, intColumnResult + i].Font.Bold = true;
                        }
                        xlWorksheet.Cells[k + 2, intColumnResult].Font.Color = System.Drawing.Color.Blue; // PASS id Blue
                        xlWorksheet.Cells[k + 2, intColumnResult + 1].Font.Color = System.Drawing.Color.Red; //FAIL is Red
                        //Set value for CELL
                        xlWorksheet.Cells[k + 2, intColumnResult] = intNumberOfPass;
                        xlWorksheet.Cells[k + 2, intColumnResult + 1] = intNumberOfFail;
                        xlWorksheet.Cells[k + 2, intColumnResult + 2] = intNumberOfNotExecute;
                    }
                }	// Exit for walkthrought all row in Test case details sheet.			    
                xlWorkBook.RefreshAll();
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
        }
        public void updateValue(string strFileName, string sheetTarget, string strValue, string strDTID, string strColumnName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook = null;
            object misValue = System.Reflection.Missing.Value;
            bool bOpenReadOnly = false;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(strFileName, 0, bOpenReadOnly, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[sheetTarget];
                int intRow = xlWorksheet.UsedRange.Rows.Count;  // Store the number of row in excel sheet
                //Get the postion of Column Name (strColumnName)
                int intColumnPosition = getPositionOfColumn(strFileName, sheetTarget, strColumnName);
                //Get the postion of Row (strDTID)
                int intRowPosition = getPositionOfRow(strFileName, sheetTarget, strDTID);
                //Set value for CELL		
                xlWorksheet.Cells[intRowPosition, intColumnPosition] = strValue;
                xlWorkBook.RefreshAll();
                xlWorkBook.Save();
                xlWorkBook.Close(false, misValue, misValue);
                ReleaseObject(xlWorksheet);
                ReleaseObject(xlWorkBook);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                log.Error(e.ToString());
            }
            finally
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
            }
        }
        //Purpose: Release used Excel Object
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                log.Error(e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
