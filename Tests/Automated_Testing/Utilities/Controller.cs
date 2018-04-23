using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using Selenium_Automated_Testing.Utilities;

//Here is the once-per-application setup information
//[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace Worksmart_Automated_Testing.Utilities
{
    public class Controller
    {
        //	Init the parameter for write log
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        readonly ReadWriteExcelToPoi _readWriteExcelToPoi = new ReadWriteExcelToPoi();
        readonly Common _common = new Common();
        readonly Keyword _keyword = new Keyword();
        readonly HtmlGenerator _htmlGenerator;
        readonly ScreenCapture _capture = new ScreenCapture();
        public PeriodicCapture PeriodCapture = new PeriodicCapture();
        public string StrBrowsertype = "";
        readonly Dictionary<string, string> _objects;

        JObject _joResultTestCaseDetails = new JObject();	// 	Store the result of each steps in suite.
        /*	joResultTestCaseDetails
        {
            "tcID":	[
                        {
                            "tsID": "<string>",
                            "result": "<boolean>"
                        }, ...,	
                    ], ..., 
        }
        */

        JObject _joResultTestSuiteDetails = new JObject();	// 	Store the result of test case.
        /*	joResultTestSuiteDetails
        {
              "tcID" <string>: "result" <boolean>, ....
        }
        */

        readonly JObject _joResultAllTestSuites = new JObject();	// 	Store the result of each test suite in TestSuites sheet
        /*	joResultAllTestSuites
        {
            "testsuiteID" : {
                                "pass": <int>,
                                "fail": "<int>,
                                "notExecute": "<int>
                            }, ...
        }
        */

        public Controller()
        {
            _objects = _common.ReadPropertiesFile(_common.OrProperties);
            _htmlGenerator = new HtmlGenerator(_common.StrHtmlReportFile);
        }

        //Purpose: This function will get Data for each steps in testcase.
        public String[] GetDataObjects(JArray jaTestSteps)
        {
            /*	jaTestSteps
            [{
                "teststepID": "<string>", 
                "description": "<string>",
                "keyword": "<string>",
                "object": "<string>",  
                "dataObject": "<string>",			
            }, ..., ]
            */
            string[] arrDataObject = null;
            try
            {
                int intLength = jaTestSteps.Count();
                int intNumberOfDataObject = 0;	//Store number of Data object
                //Get the number of Data Object
                for (int i = 0; i < intLength; i++)
                {
                    var jo = (JObject)jaTestSteps[i];
                    string strDataObject = Convert.ToString((string)jo["dataObject"]);
                    if (strDataObject != null)
                        intNumberOfDataObject++;
                }

                arrDataObject = new String[intNumberOfDataObject];

                int intIncreaseNumber = 0;
                for (int i = 0; i < intLength; i++)
                {
                    var jo = (JObject)jaTestSteps[i];
                    string strDataObject = Convert.ToString((string)jo["dataObject"]);
                    if (strDataObject != null)
                    {
                        arrDataObject[intIncreaseNumber] = strDataObject;
                        intIncreaseNumber++;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return arrDataObject;
        }

        public string GetUrl()
        {
            var strUrl = _readWriteExcelToPoi.GetUrlToExecute();
            return strUrl;
        }
        public string GetCurrentUrl(IWebDriver webDriver)
        {
            var strUrl = _readWriteExcelToPoi.GetUrlToExecute();
            var result = "";
            var url = webDriver.Url; // get the full current URL
            int startPosition = url.LastIndexOf('?') + 1;
            var wid = url.Substring(startPosition);
            result = strUrl + wid;
            return result;
        }
        //Purpose: This function will get browser type
        public string GetBrowsertype()
        {
            StrBrowsertype = _readWriteExcelToPoi.GetBrowserType().Trim();
            return StrBrowsertype;
        }
        //Purpose: This function will base on each steps in testcases and run it.
        public bool RunKeyword(IWebDriver webDriver, string strKeyword, string strObjectName, string strObjectXpath, string strData)
        {
            bool bResult = false;
            try
            {
                switch (strKeyword)
                {
                    case "NavigateUrl":
                        {
                            string strUrl = GetUrl();
                            return bResult = _keyword.NavigateUrl(webDriver, strUrl);
                        }
                    case "NavigateUrlNewPage":
                        {
                            return bResult = _keyword.NavigateUrl(webDriver, strData);
                        }
                    case "CreateNewCasePage":
                        {
                            string strUrl = GetCurrentUrl(webDriver);
                            return bResult = _keyword.NavigateUrl(webDriver, strUrl);
                        }
                    case "Input":
                        {
                            return bResult = _keyword.Input(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "Input2":
                        {
                            return bResult = _keyword.Input2(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "Click":
                        {
                            return bResult = _keyword.Click(webDriver, strObjectName, strObjectXpath);
                        }
                    case "ClickByJavaScript":
                        {
                            return bResult = _keyword.ClickByJavaScript(webDriver, strObjectName, strObjectXpath);
                        }
                    case "MouseOver":
                        {
                            return bResult = _keyword.MouseOver(webDriver, strObjectName, strObjectXpath);
                        }
                    case "GetOrderNo":
                        {
                            return bResult = _keyword.GetOrderNo(webDriver, strObjectXpath);
                        }
                    case "ClearData":
                        {
                            return bResult = _keyword.ClearData(webDriver, strObjectName, strObjectXpath);
                        }
                    case "RightClick":
                        {
                            return bResult = _keyword.RightClick(webDriver, strObjectName, strObjectXpath);
                        }
                    case "WaitTime":
                        {
                            return bResult = _keyword.WaitTime();
                        }
                    case "WaitSpecificTime":
                        {
                            return bResult = _keyword.WaitSpecificTime(strData);
                        }
                    case "Browsefile":
                        {
                            return bResult = _keyword.Browsefile(webDriver, strObjectXpath, strData);
                        }
                    case "DoubleClicks":
                        {
                            return bResult = _keyword.DoubleClicks(webDriver, strObjectName, strObjectXpath);
                        }
                    case "ScrollDownForm":
                        {
                            return bResult = _keyword.ScrollDownForm(webDriver, strObjectName, strObjectXpath);
                        }
                    case "ScrollUpForm":
                        {
                            return bResult = _keyword.ScrollUpForm(webDriver, strObjectName, strObjectXpath);
                        }
                    case "isButtonDisabled":
                        {
                            return bResult = _keyword.IsButtonDisabled(webDriver, strObjectName, strObjectXpath);
                        }
                    case "isButtonEnabled":
                        {
                            return bResult = _keyword.IsButtonEnabled(webDriver, strObjectName, strObjectXpath);
                        }
                    case "ClickByCss":
                    {
                        return bResult = _keyword.ClickByCss(webDriver, strObjectName, strObjectXpath);
                    }
                    case "SwitchWindow":
                    {
                        return bResult = _keyword.SwitchWindow(webDriver, strObjectXpath);
                    }
                    case "SwitchtoFirstWindow":
                    {
                        return bResult = _keyword.SwitchtoFirstWindow(webDriver);
                    }
                    case "SwitchtoLastWindow":
                    {
                        return bResult = _keyword.SwitchtoLastWindow(webDriver, strData);
                    }
                    case "OpenNewWindow":
                    {
                        return bResult = _keyword.OpenNewWindow(webDriver, strData);
                    }
                    case "CloseWindow":
                    {
                        return bResult = _keyword.CloseWindow(webDriver, strData);
                    }
                    case "UploadFile":
                    {
                        return bResult = _keyword.UploadFile(webDriver, strData);
                    }
                    case "SendKey":
                    {
                        return bResult = _keyword.SendKey(webDriver, strData);
                    }
                    case "MouseOverOnSubMenu":
                    {
                        return bResult = _keyword.MouseOverOnSubMenu(webDriver, strObjectName, strObjectXpath);
                    }
                    case "Quit":
                    {
                        return bResult = _keyword.Quit(webDriver);
                    }
                    case "TakeScreenShoot":
                    {
                        return bResult = _keyword.TakeScreenShoot(webDriver, strData);
                    }
                    case "ClearNewData":
                    {
                        return bResult = _keyword.ClearNewData();
                    }
                    case "ClearOldData":
                    {
                        return bResult = _keyword.ClearOldData();
                    }
                    case "CaptureOldData":
                    {
                        return bResult = _keyword.CaptureOldData(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CaptureNewData":
                    {
                        return bResult = _keyword.CaptureNewData(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CaptureOldValue":
                    {
                        return bResult = _keyword.CaptureOldValue(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CaptureNewValue":
                    {
                        return bResult = _keyword.CaptureNewValue(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CaptureOldValueOfCheckField":
                    {
                        return bResult = _keyword.CaptureOldValueOfCheckField(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CaptureNewValueOfCheckField":
                    {
                        return bResult = _keyword.CaptureNewValueOfCheckField(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifyEquals":
                    {
                        return bResult = _keyword.VerifyEquals(webDriver, strObjectName);
                    }
                    case "VerifyNotEqual":
                    {
                        return bResult = _keyword.VerifyNotEqual(webDriver, strObjectName);
                    }
                    case "VerifyGreater":
                    {
                        return bResult = _keyword.VerifyGreater(webDriver, strObjectName);
                    }
                    case "VerifyContains":
                    {
                        return bResult = _keyword.VerifyContains(webDriver);
                    }
                    case "VerifyText":
                        {
                            return bResult = _keyword.VerifyText(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "VerifyText_Contain":
                        {
                            return bResult = _keyword.VerifyText_Contain(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "VerifyText_Contain2":
                        {
                            return bResult = _keyword.VerifyText_Contain2(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "VerifyItemOfCombo":
                        {
                            return bResult = _keyword.VerifyItemOfCombo(webDriver, strObjectName, strObjectXpath, strData);
                        }
                    case "AssertUIandDatabaseFalse":
                    {
                        return bResult = _keyword.AssertUIandDatabaseFalse(strObjectName, strData);
                    }
                    case "AssertUIandDatabaseTrue":
                    {
                        return bResult = _keyword.AssertUIandDatabaseTrue(strObjectName, strData);
                    }
                    case "DismissAlert":
                    {
                        return bResult = _keyword.DismissAlert(webDriver);
                    }
                    case "AcceptAlert":
                    {
                        return bResult = _keyword.AcceptAlert(webDriver);
                    }
                    case "PressTab":
                    {
                        return bResult = _keyword.PressTab(webDriver);
                    }
                    case "PressTabOnTheObject":
                    {
                        return bResult = _keyword.PressTabOnTheObject(webDriver, strObjectName, strObjectXpath);
                    }
                    case "PressF5":
                    {
                        return bResult = _keyword.PressF5(webDriver, strObjectName, strObjectXpath);
                    }
                    case "PressEnter":
                    {
                        return bResult = _keyword.PressEnter(webDriver);
                    }
                    case "PressEnterOnTheObject":
                    {
                        return bResult = _keyword.PressEnterOnTheObject(webDriver, strObjectName, strObjectXpath);
                    }
                    case "PressArrowDown":
                    {
                        return bResult = _keyword.PressArrowDown(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CheckDataTableOfLetterTemplate":
                    {
                        return bResult = _keyword.CheckDataTableOfLetterTemplate(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "CheckColorOfAllPageButton":
                    {
                        return bResult = _keyword.CheckColorOfAllPageButton(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "CheckOrderOfDataOnOnePage":
                    {
                        return bResult = _keyword.CheckOrderOfDataOnOnePage(webDriver, strData);
                    }
                    case "IsElementReadonly":
                    {
                        return bResult = _keyword.IsElementReadonly(webDriver, strObjectName, strObjectXpath);
                    }
                    case "CheckOrderOfDataOnOneTree":
                    {
                        return bResult = _keyword.CheckOrderOfDataOnOneTree(webDriver, strData);
                    }
                    case "ClickOnItemOfCombo":
                    {
                        return bResult = _keyword.ClickOnItemOfCombo(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "ClickOnItemOfCombo2":
                    {
                        return bResult = _keyword.ClickOnItemOfCombo2(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "ClickOnItemOfDropDown":
                    {
                        return bResult = _keyword.ClickOnItemOfDropDown(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "ClickOnItemOfGroupCombo":
                    {
                        return bResult = _keyword.ClickOnItemOfGroupCombo(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "ClickOnItemOfTree":
                    {
                        return bResult = _keyword.ClickOnItemOfTree(webDriver, strData);
                    }
                    case "ClickOnAddChildItem":
                    {
                        return bResult = _keyword.ClickOnAddChildItem(webDriver, strData);
                    }
                    case "ClickOnNameOfItem":
                    {
                        return bResult = _keyword.ClickOnNameOfItem(webDriver, strData);
                    }
                    case "ClickOnNameOfItem2":
                    {
                        return bResult = _keyword.ClickOnNameOfItem2(webDriver, strData);
                    }
                    case "ClickOnCodeOfSelectActionDefinitionDialog":
                    {
                        return bResult = _keyword.ClickOnCodeOfSelectActionDefinitionDialog(webDriver, strData);
                    }
                    case "ClickOnNameOfItemOnCurrentPage":
                    {
                        return bResult = _keyword.ClickOnNameOfItemOnCurrentPage(webDriver, strData);
                    }
                    case "ClickOnNameOfItemInLetterTemplateTable":
                    {
                        return bResult = _keyword.ClickOnNameOfItemInLetterTemplateTable(webDriver, strData);
                    }
                    case "ClickOnItemOfCaseStates":
                    {
                        return bResult = _keyword.ClickOnItemOfCaseStates(webDriver, strData);
                    }
                    // use this method instead of ClickOnItemOfCaseStates, ClickOnItemOfPriorities, ClickOnItemOfActionOutcomes
                    case "ClickOnItemOfNonHierarchicalList":
                    {
                        return bResult = _keyword.ClickOnItemOfNonHierarchicalList(webDriver, strData);
                    }
                    case "ClickOnItemOfPriorities":
                    {
                        return bResult = _keyword.ClickOnItemOfPriorities(webDriver, strData);
                    }
                    case "ClickOnItemOfActionOutcomes":
                    {
                        return bResult = _keyword.ClickOnItemOfActionOutcomes(webDriver, strData);
                    }
                    case "ClickOnPageOfNoneHierarchicalList":
                    {
                        return bResult = _keyword.ClickOnPageOfNoneHierarchicalList(webDriver, strData);
                    }
                    case "ClickOnPageOfLetterTemplate":
                    {
                        return bResult = _keyword.ClickOnPageOfLetterTemplate(webDriver, strData);
                    }
                    case "ClickOnItemOfCheckBox":
                    {
                        return bResult = _keyword.ClickOnItemOfCheckBox(webDriver, strData);
                    }
                    case "UnClickOnItemOfCheckBox":
                    {
                        return bResult = _keyword.UnClickOnItemOfCheckBox(webDriver, strData);
                    }
                    case "ClickOnItemOfCheckBox2":
                    {
                        return bResult = _keyword.ClickOnItemOfCheckBox2(webDriver, strObjectName, strObjectXpath);
                    }
                    case "UnClickOnItemOfCheckBox2":
                    {
                        return bResult = _keyword.UnClickOnItemOfCheckBox2(webDriver, strObjectName, strObjectXpath);
                    }
                    case "ClickOnItemOfCheckBoxAccuracy":
                    {
                        return bResult = _keyword.ClickOnItemOfCheckBoxAccuracy(webDriver, strData);
                    }
                    case "UnClickOnItemOfSearchFieldOfCheckBox":
                    {
                        return bResult = _keyword.UnClickOnItemOfSearchFieldOfCheckBox(webDriver, strData);
                    }
                    case "ClickOnItemOfSearchFieldOfCheckBox":
                    {
                        return bResult = _keyword.ClickOnItemOfSearchFieldOfCheckBox(webDriver, strData);
                    }
                    case "UnClickOnItemOfSearchResultOfCheckBox":
                    {
                        return bResult = _keyword.UnClickOnItemOfSearchResultOfCheckBox(webDriver, strData);
                    }
                    case "ClickOnItemOfSearchResultOfCheckBox":
                    {
                        return bResult = _keyword.ClickOnItemOfSearchResultOfCheckBox(webDriver, strData);
                    }
                    case "SelectCheckBox":
                    {
                        return bResult = _keyword.SelectCheckBox(webDriver, strObjectName, strObjectXpath);
                    }
                    case "UnSelectCheckBox":
                    {
                        return bResult = _keyword.UnSelectCheckBox(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifyContentOfBreadCrumb":
                    {
                        return bResult = _keyword.VerifyContentOfBreadCrumb(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "VerifyDefaultTextInCombo":
                    {
                        return bResult = _keyword.VerifyDefaultTextInCombo(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "VerifyAnObjectPresent":
                    {
                        return bResult = _keyword.VerifyAnObjectPresent(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifyAnObjectNotPresent":
                    {
                        return bResult = _keyword.VerifyAnObjectNotPresent(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifySearchFieldNotPresent":
                    {
                        return bResult = _keyword.VerifySearchFieldNotPresent(webDriver,strData);
                    }
                    case "VerifySearchFieldPresent":
                    {
                        return bResult = _keyword.VerifySearchFieldPresent(webDriver, strData);
                    }
                    case "VerifyColorObject":
                    {
                        return bResult = _keyword.VerifyColorObject(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "VerifyBackGroundColorObject":
                    {
                        return bResult = _keyword.VerifyBackGroundColorObject(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "VerifyCheckBoxIsSelected":
                    {
                        return bResult = _keyword.VerifyCheckBoxIsSelected(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifyCheckBoxIsNotSelected":
                    {
                        return bResult = _keyword.VerifyCheckBoxIsNotSelected(webDriver, strObjectName, strObjectXpath);
                    }
                    case "VerifyUploadFileName":
                    {
                        return bResult = _keyword.VerifyUploadFileName(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "VerifyInsertIntoLetterTemplateTable":
                    {
                        return bResult = _keyword.VerifyInsertIntoLetterTemplateTable(strData);
                    }
                    case "VerifyItemOftreeIsDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOftreeIsDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOftreeIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOftreeIsNotDisplayed(webDriver, strData);
                    }
                    case "VerifyNameIsDisplayed":
                    {
                        return bResult = _keyword.VerifyNameIsDisplayed(webDriver, strData);
                    }
                    case "VerifyNameIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyNameIsNotDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOfActionOutcomesIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfActionOutcomesIsNotDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOfActionOutcomesIsDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfActionOutcomesIsDisplayed(webDriver, strData);
                    }

                    case "VerifyItemOfCaseStatesIsDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfCaseStatesIsDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOfCaseStatesIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfCaseStatesIsNotDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOfPrioritiesIsDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfPrioritiesIsDisplayed(webDriver, strData);
                    }
                    case "VerifyItemOfPrioritiesIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfPrioritiesIsNotDisplayed(webDriver, strData);
                    }
                    // use this method instead of VerifyItemOfCaseStatesIsDisplayed, VerifyItemOfPrioritiesIsDisplayed, VerifyItemOfActionOutcomesIsDisplayed
                    case "VerifyItemOfNonHierarchicalListIsDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfNonHierarchicalListIsDisplayed(webDriver, strData);
                    }
                    // use this method instead of VerifyItemOfCaseStatesIsNotDisplayed, VerifyItemOfActionOutcomesIsNotDisplayed, VerifyItemOfPrioritiesIsNotDisplayed
                    case "VerifyItemOfNonHierarchicalListIsNotDisplayed":
                    {
                        return bResult = _keyword.VerifyItemOfNonHierarchicalListIsNotDisplayed(webDriver, strData);
                    }
                    case "ClickOnItemForQuickSearch":
                    {
                        return bResult = _keyword.ClickOnItemForQuickSearch(webDriver, strData);
                    }
                    case "ClickOnItemForQuickSearchOfWorklist":
                    {
                        return bResult = _keyword.ClickOnItemForQuickSearchOfWorklist(webDriver, strData);
                    }
                    case "ClickOnCheckOfAvailableNature":
                    {
                        return bResult = _keyword.ClickOnCheckOfAvailableNature(webDriver, strData);
                    }
                    case "ClickOnLastPage":
                    {
                        return bResult = _keyword.ClickOnLastPage(webDriver, strData);
                    }
                    case "VerifyPageTitle":
                    {
                        return bResult = _keyword.VerifyPageTitle(webDriver, strObjectName, strObjectXpath, strData);
                    }
                    case "HoverMouse":
                    {
                        return bResult = _keyword.HoverMouse(webDriver, strObjectName, strObjectXpath);
                    }
                    case "GetWid":
                    {
                        return bResult = _keyword.GetWid(webDriver);
                    }
                    case "WaitAnElementPresent":
                    {
                        return bResult = _keyword.WaitAnElementPresent(webDriver, strObjectName, strObjectXpath);
                    }
                    case "WaitAnElementNotPresent":
                    {
                        return bResult = _keyword.WaitAnElementNotPresent(webDriver, strObjectName, strObjectXpath);
                    }
                    case "DeleteRule":
                    {
                        return bResult = _keyword.DeleteRule(webDriver);
                    }
                    default:
                    {
                        string strMsg = "The keyword: " + strKeyword + " does not existed in controler, please check again!!!!!!";
                        Console.WriteLine(strMsg);
                        Log.Warn(strMsg);
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return bResult;
        }
         
        //Purpose: This function will run a test cases (include all steps in that test case).
        // ReSharper disable once FunctionComplexityOverflow
        public bool RunTestCase(IWebDriver driver, string strTestcaseSheet, JObject joTestCase)
        {
            /* jaTestSteps
            [{
                "testcaseID": "<string>",
                "teststepID": "<string>",
                "description": "<string>",
                "keyword": "<string>",
                "object": "<string>",
                "dataObject": "<string>",
            }, ..., ]
            */
            bool bResultTc = true;
            JArray jaTestCase = new JArray();

            try
            {
                var strTestCaseId = (string)joTestCase["testcaseID"];
                var strDataId = (string)joTestCase["dataID"];

                //Get all steps of TC base on Test case ID then will execute
                var jaTestSteps = _readWriteExcelToPoi.GetTestSteps(strTestcaseSheet, strTestCaseId);
                /*  jaTestSteps
                [{
                    "testcaseID": "<string>",
                    "teststepID": "<string>",
                    "description": "<string>",
                    "keyword": "<string>",
                    "object": "<string>",
                    "dataObject": "<string>",
                }, ..., ]
                */
                //Get list of Data Object
                String[] arrDataObjects = GetDataObjects(jaTestSteps);
                //Get Data for TC
                var joData = _readWriteExcelToPoi.GetDataForEachTestCase(strTestcaseSheet, strDataId, jaTestSteps, arrDataObjects);
                /*  joData
                {{
                    "dataObjectName": "value", ...	  	  			
                }
                */
                int intLength = jaTestSteps.Count;
                //Walk throught all test steps and execute
                for (int i = 0; i < intLength; i++)
                {
                    JObject jo = (JObject)jaTestSteps[i];
                    string strTeststepId = (string)jo["teststepID"];
                    string strKeyword = (string)jo["keyword"];
                    strKeyword = strKeyword.Trim();
                    string strObjectName = (string)jo["object"];
                    string strObject = "";
                    JObject joTestStepsResult = new JObject();
                    /*
                    {
                        "tsID": <string>,
                        "result": "<boolean>"
                    }
                    */
                    if (!string.IsNullOrEmpty(strObjectName))
                    {
                        if (_objects.ContainsKey(strObjectName))
                        {
                            strObject = _objects[strObjectName];
                            strObject = strObject.Trim(); 
                        }
                        else
                        {
                            string strErr = "Object [" + strObjectName + "] does not existed in or.properties file, please check again";
                            Console.WriteLine(strErr);
                            Log.Error(strErr);

                            joTestStepsResult.Add("tsID", strTeststepId);
                            joTestStepsResult.Add("result", false);
                            jaTestCase.Add(joTestStepsResult);
                            bResultTc = false;
                            continue;
                        }
                    }

                    string strDataObject = (string)jo["dataObject"];
                    if (!string.IsNullOrEmpty(strDataObject))
                        strDataObject = strDataObject.Trim();
                    string strData = "";
                    if (!string.IsNullOrEmpty(strDataObject))
                        strData = (string)joData[strDataObject];
                    Console.Write("Step {0}: ", strTeststepId);
                    Log.Info(String.Format("Step {0}: ", strTeststepId));
                    //Console.WriteLine("Keyword: {0}; Object: {1}; Data: {2}", strKeyword, strObject, strData);

                    var bResultTs = RunKeyword(driver, strKeyword, strObjectName, strObject, strData);

                    joTestStepsResult.Add("tsID", strTeststepId);
                    joTestStepsResult.Add("result", bResultTs);
                    jaTestCase.Add(joTestStepsResult);
                    if (!bResultTs)
                    {
                        _capture.SaveScreenShot(strTestcaseSheet + "_" + strTeststepId + "_" + strKeyword, _common.StrHtmlReportPath);
                        bResultTc = false;
                    }
                }  // -- Exit for
                //Store the result of test steps to write the result
                /*	joResultTestCaseDetails
                {
                    "tcID":	[
                                {
                                    "tsID": <string>,
                                    "result": "<boolean>"
                                }, ...,	
                            ], ..., 
                }
                */
                _joResultTestCaseDetails.Add(strTestCaseId, jaTestCase);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }

            return bResultTc;
        }

        //Purpose: This function will get all TestSuite with Execute = Y then get All Testcase with Execute = Y to execute all of them
        public void RunTestSuite(IWebDriver driver)
        {
            /*	jaTestCaseExecute
            [{
                  "execute": "<string>",
                  "testcaseID": "<string>",
                  "description": "<string>",
                  "dataID": "<string>",
            }, ..., ]
            */
            /* jaTestSuiteExecute
            [{
                  "execute": "<string>",
                  "testsuiteID": "<string>",
                  "description": "<string>",
                  "suiteSheet": "<string>",
                  "testcaseSheet": "<string>",
            }, ..., ]
            */

            //Get list of test suites to be execute (Execute = Y)
            var jaTestSuiteExecute = _readWriteExcelToPoi.GetTestSuitesToExecute();

            try
            {

                int intNumnerOfSuite = jaTestSuiteExecute.Count;
                int intTotalNumberOfPass = 0;
                int intTotalNumberOfFail = 0;
                int intTotalNumberOfNotExecute = 0;
                _htmlGenerator.InitHtmlReport();

                //	LOOP for Run all Test Suites
                for (int i = 0; i < intNumnerOfSuite; i++)
                {
                    var joTestSuite = (JObject)jaTestSuiteExecute[i];

                    string strSuiteId = (string)joTestSuite["testsuiteID"];
                    string strDescription = (string)joTestSuite["description"];
                    string strSuiteSheet = (string)joTestSuite["suiteSheet"];
                    string strTestcaseSheet = (string)joTestSuite["testcaseSheet"];
                    String strRunningSuite = "\n===========================================================================================\n";
                    strRunningSuite = strRunningSuite + "RUNNING TEST SUITE: " + strSuiteId + " " + strDescription + "\n";
                    strRunningSuite = strRunningSuite + "===========================================================================================";
                    Console.WriteLine(strRunningSuite);
                    Log.Info(strRunningSuite);

                    //Write test suite to report
                    _htmlGenerator.CreateTableRowTestSuite(strSuiteId, strDescription);

                    //	Get list of test cases to be execute (Execute = Y)
                    var jaTestCaseExecute = _readWriteExcelToPoi.GetTestCasesToExecute(strSuiteSheet);	//Store the list of test cases will be execute

                    int intNumberOfTCs = jaTestCaseExecute.Count();
                    int intNumberOfPass = 0;
                    int intNumberOfFail = 0;

                    //// Start capturing screen.
                    //	LOOP for run all Test cases in test suite 
                    for (int j = 0; j < intNumberOfTCs; j++)
                    {
                        const string strCapture = ""; //Store the path and name of screenshoot in case fail.
                        var joTestCase = (JObject)jaTestCaseExecute[j];

                        // ReSharper disable once InconsistentNaming
                        string strTCID = (string)joTestCase["testcaseID"];
                        string strDesc = (string)joTestCase["description"];
                        // ReSharper disable once InconsistentNaming
                        string strRunningTC = "\n-------------------------------------------------------------------------------------------\n";
                        strRunningTC = strRunningTC + "Running test case: " + strTCID + " " + strDesc;
                        Console.WriteLine(strRunningTC);
                        Log.Info(strRunningTC);

                        var bResultTc = RunTestCase(driver, strTestcaseSheet, joTestCase);
                        if (bResultTc)
                        {
                            intNumberOfPass++;
                            String strResult = "The Testcase " + strTCID + " PASS";
                            Console.WriteLine(strResult);
                            Log.Info(strResult);
                        }
                        else
                        {
                            intNumberOfFail++;
                            String strResult = "The Testcase " + strTCID + " FAIL";
                            Console.WriteLine(strResult);
                            Log.Info(strResult);
                            //strCapture = screenCapture.SaveScreenShot(strTCID, common.strHTMLReportPath);
                        }
                        // Write result of test to report
                        _htmlGenerator.CreateTableRowTestCase(strTCID, strDesc, bResultTc, strCapture);
                        // 	Store the result of test case.
                        /*	joResultTestSuiteDetails
                        {
                            "tcID" <string>: "result" <boolean>, ....
                        }
                        */
                        _joResultTestSuiteDetails.Add(strTCID, bResultTc);
                    }  //	-- Exit for of run all test cases for all suite
                    int intColumnResult = _common.IntColumnOfResultTestCaseDetails;
                    // Write result to TestCase Details sheet
                    var jaTestStepsExecuted = _readWriteExcelToPoi.GetListOfTestStepsExecuted(_common.StrTestCaseFile, strSuiteSheet, strTestcaseSheet, jaTestCaseExecute);

                    string strTestCaseExecute = _readWriteExcelToPoi.GetListOfTestCasesToExecute(strSuiteSheet);
                    _readWriteExcelToPoi.WriteResultTestCaseDetails(strSuiteSheet, strTestcaseSheet, _joResultTestCaseDetails, intColumnResult, jaTestStepsExecuted, strTestCaseExecute);

                    // Write result to TestSuite Details sheet
                    int intColumnResultTestSuite = _common.IntColumnOfResultTestSuiteDetails;
                    var jaTestCases = _readWriteExcelToPoi.GetListOfTestCases(strSuiteSheet);
                    _readWriteExcelToPoi.WriteResultTestSuiteDetails(strSuiteSheet, _joResultTestSuiteDetails, intColumnResultTestSuite, jaTestCases);

                    //	Set Empty again for new suite				
                    _joResultTestCaseDetails = new JObject();
                    _joResultTestSuiteDetails = new JObject();
                    // Get the number of test case not run.

                    int intTotalTestCase = _readWriteExcelToPoi.GetNumberOfRow(_common.StrTestCaseFile, strSuiteSheet);
                    var intNumberOfNotExecute = intTotalTestCase - (intNumberOfPass + intNumberOfFail) - 1;

                    intTotalNumberOfPass += intNumberOfPass;
                    intTotalNumberOfFail += intNumberOfFail;
                    intTotalNumberOfNotExecute += intNumberOfNotExecute;

                    // Prepare data to input to final result of test suite 
                    JObject joResultTestSuite = new JObject
                    {
                        {"PASS", intNumberOfPass},
                        {"FAIL", intNumberOfFail},
                        {"notExecute", intNumberOfNotExecute}
                    };

                    //Write result of test suite to HTML file
                    _htmlGenerator.CreateTableRowResultTestSuite(joResultTestSuite);

                    /*	joResultAllTestSuites
				    {
					    "testsuiteID":	{
								  		    "pass": <int>,
								  		    "fail": "<int>,
								  		    "notExecute": "<int>
									    }
				    }
				    */
                    _joResultAllTestSuites.Add(strSuiteId, joResultTestSuite);

                }   // 	-- Exit for Run all Test Suites

                int intColumnResultPass = _common.IntColumnOfPass;
                var jaTestSuites = _readWriteExcelToPoi.GetListOfTestSuites();
                _readWriteExcelToPoi.WriteResultTestSuitesSheet(_common.StrSuiteSheet, _joResultAllTestSuites, intColumnResultPass, jaTestSuites);
                int intTotalTestCases = intTotalNumberOfPass + intTotalNumberOfFail + intTotalNumberOfNotExecute;
                JObject joTotalResultOfTest = new JObject
                {
                    {"PASS", intTotalNumberOfPass},
                    {"FAIL", intTotalNumberOfFail},
                    {"notExecute", intTotalNumberOfNotExecute},
                    {"totalTestcase", intTotalTestCases}
                };
                //Insert the Final result to HTML report.
                _htmlGenerator.CreateTableRowFinalResult(joTotalResultOfTest);
                //Insert Actual End time for test run.
                _htmlGenerator.InsertEndTime();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
        }
    }
}
