using System;
using System.IO;
using System.Reflection;
using log4net;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;

namespace Selenium_Automated_Testing.Utilities
{
    public class ReadWriteExcelToPoi
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private readonly Common _common = new Common();
        public ReadWriteExcelToPoi()
        {
            CopyFile();
        }
         
        //Purpose: This function will Copy TestCase File to Write the Result to it
        private void CopyFile()
        {
            const bool bOverWritten = true;
            try
            {
                File.Copy(_common.StrTestCaseFile, _common.StrResultFile, bOverWritten);
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }

        // Purpose: This function will read all rows of specific sheet with limit of column
        public string[,] ReadInputData(string strFileName, string strSheetName, int intColumn)
        {
            string[,] arrData = null;
            try
            {
                using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    if (sheet !=null)
                    {
                        int countRow = sheet.LastRowNum;
                        arrData = new string[countRow, intColumn];
                        int intR = 0;
                        int intC = 0;
                        for (int i = 1; i <= sheet.LastRowNum; i++, intR++)
                        {
                            IRow row = sheet.GetRow(i);
                            for (int j = 0; j < intColumn; j++, intC++)
                            {
                                ICell cell = row.GetCell(j);
                                if (cell != null)
                                {
                                    arrData[intR, intC] = ValueCell(cell);
                                }
                                else
                                {
                                    arrData[intR, intC] = "";
                                }
                            }
                            intC = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.ToString());
            }

            return arrData;
        }
        // This function will read all rows of specific sheet with all column
        public string[,] ReadInputData(string strFileName, string strSheetName)
        {

            string[,] arrData = null;
            try
            {
                using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    if (sheet != null)
                    { 
                        int countRow = sheet.LastRowNum + 1;
                        int countColumn = sheet.GetRow(sheet.FirstRowNum).LastCellNum;
                        arrData = new string[countRow, countColumn];
                        int intR = 0;
                        int intC = 0;
                        for (int i = 0; i < countRow; i++, intR++)
                        {
                            IRow row = sheet.GetRow(i);
                            for (int j = 0; j < countColumn; j++, intC++)
                            {
                                ICell cell = row.GetCell(j);
                                if (cell != null)
                                {
                                    var strTemp = ValueCell(cell);
                                    arrData[intR, intC] = strTemp;
                                }
                                else
                                {
                                    arrData[intR, intC] = "";
                                }
                            }
                            intC = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine("{0} Exception caught.", ex);
                Log.Error(ex.ToString());
            }
            return arrData;
        }
        // Purpose: This function will get the position of COLUMN base on specific value and sheet name
        public int GetPositionOfColumn(string strFileName, string strSheetName, string strValue)
        {
            int positionColumn = 99999;
            try
            {
                using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                if (ValueCell(cell).Equals(strValue.Trim()))
                                {
                                    positionColumn = j + 1;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                Log.Error(ex.ToString());
            }
            return positionColumn;
        }
        //Purpose: This function will get the position of ROW base on specific value and sheet name
        public int GetPositionOfRow(string strFileName, string strSheetName, string strValue)
        {
            int positionRow = 99999;
            try
            {
                using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    for (int i = 0; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                if (ValueCell(cell).Equals(strValue.Trim()))
                                {
                                    positionRow = i + 1;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                Log.Error(ex.ToString());
            }
            return positionRow;
        }
        public JArray GetListOfTestSteps(string testcasePath, String strTestCaseSheet)
        {
            JArray jaTestSteps = new JArray();
            try
            {
                using (FileStream file = new FileStream(testcasePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strTestCaseSheet);
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        /*	(0)TestCaseID	(1)StepID	(2)Description
                            Here we just get TestcaseID and StepID
                        */
                        IRow row = sheet.GetRow(i);
                        string strTcid = Convert.ToString(ValueCell(row.Cells[0]));
                        string strTsid = Convert.ToString(ValueCell(row.Cells[1]));
                        JObject joTestStep = new JObject();
                        if (!string.IsNullOrEmpty(strTcid))
                        {
                            joTestStep.Add("testcaseID", strTcid);
                            joTestStep.Add("teststepID", strTsid);
                            jaTestSteps.Add(joTestStep);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.ToString());
            }
            return jaTestSteps;
        }
        public JArray GetListOfTestStepsExecuted(string strTestCasePath, String strSuiteSheet, String strTestCaseSheet, JArray jaTestCaseExecute)
        {
            JArray jaTestSteps = new JArray();
            try
            {
                using (FileStream file = new FileStream(strTestCasePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strTestCaseSheet);
                    // PROCESS EACH SPREADSHHET ROW      
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        /*	(0)TestCaseID	(1)StepID	(2)Description
				 	    Here we just get TestcaseID and StepID
    				    */
                        IRow row = sheet.GetRow(i);
                        string strTestCaseId = ValueCell(row.GetCell(0));
                        for (int j = 1; j <= jaTestCaseExecute.Count; j++)
                        {
                            JObject joTestCaseExecuted = (JObject)jaTestCaseExecute[j - 1];
                            string strTestCaseExecuted = (string)(joTestCaseExecuted["testcaseID"]);

                            if (strTestCaseId.Equals(strTestCaseExecuted))
                            {
                                string strTcid = ValueCell(row.GetCell(0));
                                string strTsid = ValueCell(row.GetCell(1));
                                // ReSharper disable once UseObjectOrCollectionInitializer
                                JObject joTestStep = new JObject();
                                joTestStep.Add("testcaseID", strTcid);
                                joTestStep.Add("teststepID", strTsid);
                                jaTestSteps.Add(joTestStep);
                                break;	
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.ToString());
            }
            return jaTestSteps;
        }
        public int GetNumberOfColumn(string strFilePath, string strSheetName)
        {
            int intColumn = 0;
            try
            {
                using (FileStream file = new FileStream(strFilePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    if (sheet != null)
                    {
                        intColumn = sheet.GetRow(sheet.FirstRowNum).LastCellNum;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return intColumn;
        }
        public int GetNumberOfRow(string strFilePath, string strSheetName)
        {
            int intRow = 0;
            try
            {
                using (FileStream file = new FileStream(strFilePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSheetName);
                    intRow = sheet.LastRowNum + 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return intRow;
        }
        public void WriteResultTestCaseDetails(String strSuiteSheet, String strTestCaseSheet, JObject joResultTestCaseDetails, int intColumnResult, JArray jaTestCaseExecute, string strTestCaseExecute)
        {
            try
            {
                Common common = new Common();
                using (FileStream file = new FileStream(common.StrResultFile, FileMode.OpenOrCreate))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strTestCaseSheet);
                    int intRow = sheet.LastRowNum + 1;
                    var jaTestSteps = GetListOfTestSteps(common.StrResultFile, strTestCaseSheet);	// This will store the list all TCs in TEST CASE DETAILS sheet
                    var jaTestStepsExecuted = GetListOfTestStepsExecuted(common.StrResultFile, strSuiteSheet, strTestCaseSheet, jaTestCaseExecute);	// This will store the list TC that execute = Y

                    for (int k = 0; k < intRow - 1; k++)
                    {
                        JObject joTestCase = (JObject)jaTestSteps[k];
                        string strTcid = (string)joTestCase["testcaseID"];
                        string strTsid = (string)joTestCase["teststepID"];
                        if (strTestCaseExecute.Contains(strTcid))
                        {
                            int intLenght = jaTestStepsExecuted.Count;
                            //Go through all row of test details that had been executed. 
                            for (int i = 0; i < intLenght; i++)
                            {
                                JObject joTestCaseExecuted = (JObject)jaTestStepsExecuted[i];
                                string strTcidExecuted = (string)joTestCaseExecuted["testcaseID"];
                                string strTsidExecuted = (string)joTestCaseExecuted["teststepID"];
                                if (strTcid.Equals(strTcidExecuted) && strTsid.Equals(strTsidExecuted))
                                {
                                    var jaTestStepsResult = (JArray)joResultTestCaseDetails[strTcidExecuted];
                                    /*[
                                        {
                                            "tsID": <string>,
                                            "result": "<boolean>"
                                        }, ......,	
                                    ]*/
                                    int intLenghtOfTc = jaTestStepsResult.Count;
                                    //Go through all row of test details.
                                    for (int j = 0; j < intLenghtOfTc; j++)
                                    {
                                        JObject joTestStepResult = (JObject)jaTestStepsResult[j];
                                        string strTsidActual = (string)joTestStepResult["tsID"];
                                        bool strResult = (bool)joTestStepResult["result"];

                                        if (strTsidExecuted.Equals(strTsidActual))	// Test step match with result
                                        {
                                            //Set Format for FONT    
                                            IFont fontcell = hssfworkbook.CreateFont();
                                            fontcell.Boldweight = (short)FontBoldWeight.Bold;

                                            string strResultToWrite;
                                            if (strResult)
                                            {
                                                fontcell.Color = HSSFColor.Blue.Index;
                                                strResultToWrite = "PASS";
                                            }
                                            else
                                            {
                                                fontcell.Color = HSSFColor.Red.Index;
                                                strResultToWrite = "FAIL";
                                            }
                                            //Set style for CELL                                        
                                            ICellStyle style = hssfworkbook.CreateCellStyle();
                                            style.SetFont(fontcell);
                                            style.BorderLeft = BorderStyle.Thin;
                                            style.BorderRight = BorderStyle.Thin;
                                            style.BorderTop = BorderStyle.Thin;
                                            style.BorderBottom = BorderStyle.Thin;
                                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).CellStyle = style;
                                            //Set value for CELL
                                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).SetCellValue(strResultToWrite);
                                            break;
                                        }	// End if
                                    }	// Exit For
                                    break;
                                }	// End if
                            }	//	Exit for walkthrought all test case that executed	
                        }	// End if
                    }
                    WriteDataToExcel(common.StrResultFile, hssfworkbook);
                    // Exit for walkthrought all row in Test case details sheet.
                }
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }
        // purpose: Write result to TestSuiteDetails
        /// <param name="strSuiteSheet">Name Suite Sheet</param>
        /// <param name="joResultTestSuiteDetails"></param>
        /// <param name="intColumnResult">get the position of Column result in test suite details sheet of test case file </param>
        /// <param name="jaTestCases">This will store the list all TCs in TEST SUITE DETAILS sheet</param>
        public void WriteResultTestSuiteDetails(String strSuiteSheet, JObject joResultTestSuiteDetails, int intColumnResult, JArray jaTestCases)
        {
            try
            {
                /*	joResultTestSuiteDetails
                {
                    "tcID" <string>: "result" <boolean>, ....
                }
            */
                Common common = new Common();
                using (FileStream file = new FileStream(common.StrResultFile, FileMode.OpenOrCreate))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strSuiteSheet);
                    int intRow = sheet.LastRowNum;
                    for (int k = 0; k < intRow; k++)
                    {
                        JObject joTestCase = (JObject)jaTestCases[k];
                        string strExecute = (string)joTestCase["execute"];
                        string strTcid = (string)joTestCase["testcaseID"];
                        if (strExecute.Equals("Y"))
                        {
                            string strResultToWrite;
                            bool strResult = (bool)joResultTestSuiteDetails[strTcid];
                            //Set Format for FONT 
                            IFont fontcell = hssfworkbook.CreateFont();
                            fontcell.Boldweight = (short)FontBoldWeight.Bold;
                            if (strResult)
                            {
                                fontcell.Color = HSSFColor.Blue.Index;
                                strResultToWrite = "PASS";
                            }
                            else
                            {
                                fontcell.Color = HSSFColor.Red.Index;
                                strResultToWrite = "FAIL";
                            }
                            //Set format for cell
                            ICellStyle style = hssfworkbook.CreateCellStyle();
                            style.SetFont(fontcell);
                            style.BorderLeft = BorderStyle.Thin;
                            style.BorderRight = BorderStyle.Thin;
                            style.BorderTop = BorderStyle.Thin;
                            style.BorderBottom = BorderStyle.Thin;
                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).CellStyle = style;
                            //Set value for CELL
                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).SetCellValue(strResultToWrite);
                        }
                    }
                    WriteDataToExcel(common.StrResultFile, hssfworkbook);
                    // Exit for walkthrought all row in Test case details sheet.
                }
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }
        //Purpose: Write result to TestSuites Sheet
        public void WriteResultTestSuitesSheet(String strTestSuitesSheet, JObject joResultAllTestSuites, int intColumnResult, JArray jaTestSuites)
        {
            try
            {
                Common common = new Common();
                using (FileStream file = new FileStream(common.StrResultFile, FileMode.OpenOrCreate))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                    ISheet sheet = hssfworkbook.GetSheet(strTestSuitesSheet);
                    int intRow = sheet.LastRowNum;
                    //Go through all row of testsuites sheet
                    for (int k = 0; k < intRow; k++)
                    {
                        JObject joTestCase = (JObject)jaTestSuites[k];
                        string strExecute = (string)joTestCase["execute"];
                        string strTestSuiteId = (string)joTestCase["testsuiteID"];

                        if (strExecute.Equals("Y"))
                        {
                            /*{
                                "pass": <int>,
                                "fail": "<int>,
                                "notExecute": "<int>
                            }, ...*/
                            var joTestSuiteResult = (JObject)joResultAllTestSuites[strTestSuiteId];
                            int intNumberOfPass = (int)joTestSuiteResult["PASS"];
                            int intNumberOfFail = (int)joTestSuiteResult["FAIL"];
                            int intNumberOfNotExecute = (int)joTestSuiteResult["notExecute"];

                            //Set Format for FONT
                            for (int i = 0; i < 2; i++)
                            {
                                IFont fontcelli = hssfworkbook.CreateFont();
                                fontcelli.Boldweight = (short)FontBoldWeight.Bold;
                                ICellStyle stylei = hssfworkbook.CreateCellStyle();
                                stylei.BorderLeft = BorderStyle.Thin;
                                stylei.BorderRight = BorderStyle.Thin;
                                stylei.BorderTop = BorderStyle.Thin;
                                stylei.BorderBottom = BorderStyle.Thin;
                                stylei.SetFont(fontcelli);
                                sheet.GetRow(k + 1).GetCell(intColumnResult + i).CellStyle = stylei;
                            }
                            //Set format font for cell
                            IFont fontcellTrue = hssfworkbook.CreateFont();
                            fontcellTrue.Boldweight = (short)FontBoldWeight.Bold;
                            fontcellTrue.Color = HSSFColor.Blue.Index;
                            ICellStyle styleTrue = hssfworkbook.CreateCellStyle();
                            styleTrue.SetFont(fontcellTrue);
                            styleTrue.BorderLeft = BorderStyle.Thin;
                            styleTrue.BorderRight = BorderStyle.Thin;
                            styleTrue.BorderTop = BorderStyle.Thin;
                            styleTrue.BorderBottom = BorderStyle.Thin;
                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).CellStyle = styleTrue;

                            IFont fontcellFalse = hssfworkbook.CreateFont();
                            fontcellFalse.Boldweight = (short)FontBoldWeight.Bold;
                            fontcellFalse.Color = HSSFColor.Red.Index;
                            ICellStyle styleFalse = hssfworkbook.CreateCellStyle();
                            styleFalse.SetFont(fontcellFalse);
                            styleFalse.BorderLeft = BorderStyle.Thin;
                            styleFalse.BorderRight = BorderStyle.Thin;
                            styleFalse.BorderTop = BorderStyle.Thin;
                            styleFalse.BorderBottom = BorderStyle.Thin;
                            sheet.GetRow(k + 1).GetCell(intColumnResult).CellStyle = styleFalse;
                            //Set value for CELL
                            sheet.GetRow(k + 1).GetCell(intColumnResult - 1).SetCellValue(intNumberOfPass);
                            sheet.GetRow(k + 1).GetCell(intColumnResult).SetCellValue(intNumberOfFail);
                            sheet.GetRow(k + 1).GetCell(intColumnResult + 1).SetCellValue(intNumberOfNotExecute);
                        }
                    }	// Exit for walkthrought all row in Test case details sheet.
                    WriteDataToExcel(common.StrResultFile, hssfworkbook);
                }
            }
            catch (Exception e)
            {
                Log.Error(e.ToString());
            }
        }
        private void WriteDataToExcel(string strPathfile, HSSFWorkbook hssfworkbook)
        {
            FileStream fileUpdate = File.Create(strPathfile);
            hssfworkbook.Write(fileUpdate);
            fileUpdate.Close();//
        }
        //Purpose: Get value of cell
        public string ValueCell(ICell cell)
        {
            string strValue;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    strValue = null;
                    break;
                case CellType.Boolean:
                    strValue = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Numeric:
                    strValue = cell.ToString();    //This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number.
                    break;
                case CellType.String:
                    strValue = cell.StringCellValue;
                    break;
                case CellType.Error:
                    strValue = cell.ErrorCellValue.ToString();
                    break;
                default:
                    strValue = "=" + cell.CellFormula;
                    break;
            }
            return strValue;
        }
        //Purpose: This function will get only filename in fullpath of file
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
                Log.Error(e.ToString());
            }
            return strOnlyFileName;
        }
         
        //Purpose: This function will get all TestSuites that Execute = Y then return.
        public JArray GetTestSuitesToExecute()
        {
            const int intColumn = 5; //(0)Execute		(1)TestSuiteID		(2)Description	(3)SuiteSheet	(4)TestCaseSheet            
            string[,] arrData = ReadInputData(_common.StrTestCaseFile, _common.StrSuiteSheet, intColumn);

            JArray jaTestSuitesExecute = new JArray();  // This to store list of Test Suites will be execute.          
            /* jaTestSuiteExecute
            [{
                "execute": "<string>",
                "testsuiteID": "<string>",
                "description": "<string>",
                "suiteSheet": "<string>",
                "testcaseSheet": "<string>",
            }, ..., ]
            */
            try
            {
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseId = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strSuiteSheet = arrData[intR, 3];
                    string strTestCaseSheet = arrData[intR, 4];
                    if (strExecute.Equals("Y"))
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        JObject joTestSuite = new JObject();
                        joTestSuite.Add("execute", strExecute);
                        joTestSuite.Add("testsuiteID", strTestCaseId);
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
                Log.Error(e.ToString());
            }
            return jaTestSuitesExecute;
        }
        //Purpose: This function will get all TestSuites in TestSuite sheet then return the list of TestSuites
        public JArray GetListOfTestSuites()
        {
            const int intColumn = 2; //(0)Execute		(1)TestSuiteID
            string[,] arrData = ReadInputData(_common.StrTestCaseFile, _common.StrSuiteSheet, intColumn);//ReadInputData(strTestCaseFile, common.strSuiteSheet, intColumn);

            JArray jaTestSuitesExecute = new JArray();  // This to store list of Test Suites will be execute.          
            /* jaTestSuiteExecute
            [{
                "execute": "<string>",
                "testsuiteID": "<string>",            
            }, ..., ]
            */
            try
            {
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestSuiteId = arrData[intR, 1];
                    if (!string.IsNullOrEmpty(strTestSuiteId))
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        JObject joTestSuite = new JObject();
                        joTestSuite.Add("execute", strExecute);
                        joTestSuite.Add("testsuiteID", strTestSuiteId);
                        jaTestSuitesExecute.Add(joTestSuite);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }

            return jaTestSuitesExecute;
        }
         
        //Purpose: This function will get all TestCases that Execute = Y then return it in JArray
        public JArray GetTestCasesToExecute(string strSuiteSheet)
        {
            JArray jaTestCasesExecute = new JArray();	//Store the list of test cases will be execute
            /*	jaTestCaseExecute
            [{
                  "execute": "<string>",
                  "testcaseID": "<string>",
                  "description": "<string>",
                  "dataID": "<string>",
            }, ..., ]
            */
            const int intColumn = 4; //(0)Execute		(1)TestCase		(2)Description	(3)DataID

            try
            {
                string[,] arrData = ReadInputData(_common.StrTestCaseFile, strSuiteSheet, intColumn);
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseId = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strDataId = arrData[intR, 3];
                    if (strExecute.Equals("Y"))
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        JObject joInputData = new JObject();
                        joInputData.Add("execute", strExecute);
                        joInputData.Add("testcaseID", strTestCaseId);
                        joInputData.Add("description", strDescription);
                        joInputData.Add("dataID", strDataId);
                        jaTestCasesExecute.Add(joInputData);
                    }
                }
            }
            catch (Exception e)
            {
                //Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
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
            const int intColumn = 2; //(0)Execute		(1)TestCase
            try
            {
                string[,] arrData = ReadInputData(_common.StrTestCaseFile, strSuiteSheet, intColumn);

                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseId = arrData[intR, 1];
                    if (!string.IsNullOrEmpty(strTestCaseId))
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        JObject joInputData = new JObject();
                        joInputData.Add("execute", strExecute);
                        joInputData.Add("testcaseID", strTestCaseId);
                        jaTestCasesExecute.Add(joInputData);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }

            return jaTestCasesExecute;
        }
        //Purpose: This function will get all TestCases that Execute = Y then return it in a string with only TestCaseID in it.
        public string GetListOfTestCasesToExecute(string strSuiteSheet)
        {
            string strTestCaseExecute = "";
            JArray jaExecuteTestcases = new JArray();	//["testcaseID 001", "testcaseID 002" "testcaseID 003", ....]
            // ReSharper disable once ConvertToConstant.Local
            int intColumn = 2; //(0)Execute		(1)TestCase
            try
            {
                string[,] arrData = ReadInputData(_common.StrTestCaseFile, strSuiteSheet, intColumn);
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strExecute = arrData[intR, 0];
                    string strTestCaseId = arrData[intR, 1];
                    if (strExecute.Equals("Y"))
                    {
                        jaExecuteTestcases.Add(strTestCaseId);
                    }
                }
                strTestCaseExecute = jaExecuteTestcases.ToString();
            }
            catch (Exception e)
            {
                //Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
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
            const int intColumn = 6; //(0)TestCaseID	(1)StepID	(2)Description	(3)Keyword	(4)Object	(5)DataObject

            try
            {
                string[,] arrData = ReadInputData(_common.StrTestCaseFile, strSheetName, intColumn);
                int intRow = arrData.GetLength(0);
                for (int intR = 0; intR < intRow; intR++)
                {
                    string strTestCaseId = arrData[intR, 0];
                    string strTestStepId = arrData[intR, 1];
                    string strDescription = arrData[intR, 2];
                    string strKeyword = arrData[intR, 3];
                    string strObject = arrData[intR, 4];
                    string strDataObject = arrData[intR, 5];
                    if (strTestCaseExecute.Equals(strTestCaseId))
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        JObject joTestcase = new JObject();
                        joTestcase.Add("testcaseID", strTestCaseId);
                        joTestcase.Add("teststepID", strTestStepId);
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
                Log.Error(e.ToString());
            }
            return jaExecuteSteps;
        }
        //Purpose: get data for test case
        public JObject GetDataForEachTestCase(String strSheetName, String strDataId, JArray jaTestSteps, String[] arrDataObjects)
        {
            JObject joExecuteData = new JObject();
            /*
            {
                "dataObjectName" : "value",...	  	  			
            }
            */
            try
            {
                string[,] arrData = ReadInputData(_common.StrDataFile, strSheetName);
                int intLength = arrDataObjects.Length;
                int intNumberOfColumnData = GetNumberOfColumn(_common.StrDataFile, strSheetName);
                int intIndexOfRow = 0; //Store the index of row data
                //Get index of ROW to get data
                if (arrData != null)
                {
                    for (int i = 0; i < arrData.GetLength(0); i++)
                    {
                        if (strDataId.Equals(arrData[i, 0]))
                        {
                            intIndexOfRow = i;
                            break;
                        }
                    }
                }
                //Get Data and input to JSON Object
                for (int i = 0; i < intLength; i++)
                {
                    string strDataObjectName = arrDataObjects[i];
                    for (int j = 0; j < intNumberOfColumnData; j++)
                    {
                        string strHeader = "";
                        if (arrData[0, j] != null)
                        {
                            strHeader = arrData[0, j].ToString();
                            strHeader = strHeader.Trim();
                        }
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
                Log.Error(e.ToString());
            }
            return joExecuteData;
        }
        //Purpose: Get URL to execute.
        public string GetUrlToExecute()
        {
            const int intColumn = 3; //(0)Execute		(1)TestSuiteID		(2)Description	(3)SuiteSheet	(4)TestCaseSheet            
            string[,] arrData = ReadInputData(_common.StrDataFile, _common.StrSheetNameUrl, intColumn);
            string strUrl = "";

            for (int i = 0; i < arrData.Length; i++)
            {
                if (i == 0)
                {
                    strUrl = arrData[0, i];
                }
            }
            return strUrl;
        }
        //Purpose: Read Data xls file to get the browser type
        public string GetBrowserType()
        {
            const int intColumn = 3;
            string[,] arrData = ReadInputData(_common.StrDataFile, _common.StrSheetNameUrl, intColumn);
            string strBrowserType = "";

            for (int i = 2; i < arrData.Length; i++)
            {
                if (i == 2)
                {
                    strBrowserType = arrData[0, i];
                }
            }
            return strBrowserType;
        }
    }
}
