using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using AutoItX3Lib;
using Castle.Core.Internal;
using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using Selenium_Automated_Testing.Utilities;
//using System.Windows.Forms;
namespace Worksmart_Automated_Testing.Utilities
{
    public partial class Keyword
    {
        //	Init the parameter for write log
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private readonly Common _clsCommon = new Common();
        private string _strOderNo = string.Empty;
        private readonly AutoItX3 _autoit = new AutoItX3();
        private readonly ScreenCapture _capture = new ScreenCapture();
        /// <summary>
        ///TakeScreenShoot() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: Take a photo of screen
        /// </summary>
        public bool TakeScreenShoot(IWebDriver webDriver, String strFileName)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                bResult = true;
                const string strStepDetails = "Take the screenshoot successfully";
                _capture.SaveScreenShot(strFileName, _clsCommon.StrHtmlReportPath);
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///Quit() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: Close browser and stop execution
        /// </summary>
        public bool Quit(IWebDriver webDriver)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                webDriver.Quit();
                bResult= true;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///MouseOverOnSubMenu() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: click on a hiden sub menu
        ///Description: use JQuery command and execute by JavaScript command
        /// </summary>
        public bool MouseOverOnSubMenu(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            int i = 0;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    if (js != null) js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                    strStepDetails.AppendFormat("Passed. Mouse over on {0} successfully.",strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fai. The {0} is NOT displayed.",strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///SendKey(IWebDriver webDriver, string strData)
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: send key from keyboard, like that: Enter(Return)...
        ///Description: use AutoIT library
        /// </summary>
        public bool SendKey(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                _autoit.Send(strData);
                bResult = true;
                var strStepDetails = "Send key [" + strData + "] successfully";
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///AttachFile()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: Upload file 
        ///Description: use AutoIT library 
        /// </summary>
        public bool UploadFile(IWebDriver webDriver, String strData)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                _autoit.WinWaitActive("Open");
                _autoit.Send(strData);
                _autoit.Send("{ENTER}");
                //SendKeys.SendWait(@strData);
                //SendKeys.SendWait(@"{Enter}");
                bResult = true;
                const string strStepDetails = "Passed. Attach the file successfully";
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///AttachFile()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: Upload file 
        ///Description: use AutoIT library 
        /// </summary>
        public bool UploadFileOnBrowserStack(IWebDriver webDriver, String strData)
        {
            bool bResult = false;
            //Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                _autoit.WinWaitActive("Open");
                _autoit.Send(strData);
                _autoit.Send("{ENTER}");
                //SendKeys.SendWait(@strData);
                //SendKeys.SendWait(@"{Enter}");
                bResult = true;
                const string strStepDetails = "Passed. Attach the file successfully";
                Console.WriteLine(strStepDetails);
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }

        /// <summary>
        ///CloseWindow()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Now-2015
        ///Purpose: Move to new tab or page or popup 
        ///Description: use action SwitchTo of Selenium
        /// </summary>
        /// <Parameter>
        /// strObjectXpath:is title of new tab (or page, popup...)
        /// </Parameter>
        public bool CloseWindow(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var script = string.Format("window.close('{0}', '_blank');", strData);
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                if (js != null) js.ExecuteScript(script);
                bResult = true;
                const string strStepDetails = "Passed. Close the window successfully";
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///OpenNewWindow()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Now-2015
        ///Purpose: Move to new tab or page or popup 
        ///Description: use action SwitchTo of Selenium
        /// </summary>
        /// <Parameter>
        /// strObjectXpath:is title of new tab (or page, popup...)
        /// </Parameter>
        public bool OpenNewWindow(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var script = string.Format("window.open('{0}', '_blank');", strData);
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            String useragent = (String)js.ExecuteScript("return navigator.userAgent;");
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                if (!useragent.Contains("MSIE") && !useragent.Contains("rv:11.0"))
                {
                    js.ExecuteScript(script);
                }
                else
                {
                    webDriver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
                    //webDriver.Navigate().GoToUrl(strData);
                    //webDriver.Navigate().Forward();
                    //_clsCommon.NavigateUrl(webDriver, strData);

                    //IJavaScriptExecutor jscript = webDriver as IJavaScriptExecutor;
                    //jscript.ExecuteScript("window.open()");
                    //List<string> handles = webDriver.WindowHandles.ToList<string>();
                    //webDriver.SwitchTo().Window(handles.Last());
                }
                bResult = true;
                const string strStepDetails = "Passed. Navigate URL successfully";
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///SwitchWindow()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2014
        ///Purpose: Move to new tab or page or popup 
        ///Description: use action SwitchTo of Selenium
        /// </summary>
        /// <Parameter>
        /// strObjectXpath:is title of new tab (or page, popup...)
        /// </Parameter>
        public bool SwitchWindow(IWebDriver webDriver, string strObjectXpath)
        {
            var currentWindow = webDriver.CurrentWindowHandle;
            var availableWindows = new List<string>(webDriver.WindowHandles);
            Thread.Sleep(_clsCommon.IntTimeWait);
            foreach (string w in availableWindows)
            {
                if (w != currentWindow)
                {
                    webDriver.SwitchTo().Window(w);
                    if (webDriver.Title == strObjectXpath)
                    {
                        return true;
                    }
                    webDriver.SwitchTo().Window(currentWindow);
                }
            }
            return false;
        }
        /// <summary>
        ///SwitchtoFirstWindow()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Nov-2015
        ///Purpose: Move to the last tab or page  
        ///Description: use action SwitchTo() of Selenium
        /// </summary>
        public bool SwitchtoFirstWindow(IWebDriver webDriver)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                webDriver.SwitchTo().Window(webDriver.WindowHandles.First());
                bResult = true;
                const string strStepDetails = "Passed. Switch to first window successfully";
                Console.WriteLine(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        ///SwitchtoLastWindow()
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Nov-2015
        ///Purpose: Move to the last tab or page  
        ///Description: use action SwitchTo() of Selenium
        /// </summary>
        public bool SwitchtoLastWindow(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                webDriver.SwitchTo().Window(webDriver.WindowHandles.Last());
                _clsCommon.NavigateUrl(webDriver, strData);
                bResult = true;
                const string strStepDetails = "Switch to last window successfully";
                Console.WriteLine(strStepDetails);
            }
             catch (Exception ex)
             {
                 Console.WriteLine(ex);
                 Log.Error(ex.ToString());
             }
             return bResult;
        }
        /// <summary>
        /// Purpose: This function helps to navigate to a URL which was set in TestData excel file, General sheet, URL column
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="strUrl">URL was set in TestData excel file, General sheet, URL column</param>
        /// <returns>
        ///     True: if navigate to URL successfully
        ///     False: if navigate to URL unsuccessfully
        /// </returns>

        public bool NavigateUrl(IWebDriver webDriver, string strUrl)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                _clsCommon.NavigateUrl(webDriver, strUrl);
                bResult = true;
                const string strStepDetails = "Passed. Navigate URL successfully";
                Thread.Sleep(_clsCommon.IntClickWait);
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        /// Purpose: This function helps to select a file
        /// </summary>
        public bool Browsefile(IWebDriver webDriver, string oObjectXpath, string strData)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                IWebElement fileInput = webDriver.FindElement(By.XPath(oObjectXpath));
                fileInput.SendKeys(strData);
                const string strStepDetails = "Browse the file successfully";
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        /// This function helps to input data to textbox object
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <param name="strData">Data to fill in object</param>
        /// <returns>
        ///     True: if input data to textbox successfully
        ///     False: if input data to textbox unsuccessfully
        /// </returns>
        public bool Input(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                ReadOnlyCollection<IWebElement> oWebElement;
                do
                {
                    oWebElement = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElement.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElement.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElement[0]).Build().Perform();
                    if ((strData.IsNullOrEmpty() || strData == "") && (_strOderNo.IsNullOrEmpty() || _strOderNo == ""))
                    {
                        strStepDetails.AppendFormat("Fail. There isn't any Test Data.");
                    }
                    //Input data for searching function when strData is blank
                    else
                    {
                        if (strData.IsNullOrEmpty() || strData == "")
                        {
                            oWebElement[0].SendKeys(_strOderNo.Trim());
                            strStepDetails.AppendFormat("Passed.Enter {0} into {1}", strData, strObjectName);
                            bResult = true;
                        }
                        //Input data for both searching function and other textboxes
                        else
                        {
                            oWebElement[0].Clear();
                            oWebElement[0].SendKeys(strData);
                            strStepDetails.AppendFormat("Passed.Enter {0} into {1}", strData, strObjectName);
                            bResult = true;
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        public bool Input2(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElement;
                do
                {
                    oWebElement = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElement.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElement.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElement[0]).Build().Perform();
                    if ((strData.IsNullOrEmpty() || strData == "") && (_strOderNo.IsNullOrEmpty() || _strOderNo == ""))
                    {
                        strStepDetails.AppendFormat("Fail. There isn't any Test Data.");
                    }
                    //Input data for searching function when strData is blank
                    else
                    {
                        if (strData.IsNullOrEmpty() || strData == "")
                        {
                            oWebElement[0].SendKeys(_strOderNo.Trim());
                            strStepDetails.AppendFormat("Passed.Enter {0} into {1}", strData, strObjectName);
                            bResult = true;
                        }
                        //Input data for both searching function and other textboxes
                        else
                        {
                            oWebElement[0].Clear();
                            oWebElement[0].SendKeys(strData);
                            strStepDetails.AppendFormat("Passed.Enter {0} into {1}", strData, strObjectName);
                            bResult = true;
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        
        /// <summary>
        /// This function helps to click on an object
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if click on object successfully
        ///     False: if click on object unsuccessfully
        /// </returns>
        public bool Click(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                ReadOnlyCollection<IWebElement> webObject;
                do
                {
                    webObject = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!webObject.Any() && i < _clsCommon.IntTimeSearch);
                if (webObject.Any())
                {
                    Actions builder = new Actions(webDriver);
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    builder.MoveToElement(webObject[0]).Build().Perform();
                    builder.Click(webObject[0]).Build().Perform();

                    strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        public bool ClickByJavaScript(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> webObject;
                do
                {
                    webObject = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!webObject.Any() && i < _clsCommon.IntTimeSearch);
                if (webObject.Any())
                {
                    Actions builder = new Actions(webDriver);
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    ((IJavaScriptExecutor)webDriver).ExecuteScript("arguments[0].fireEvent('onclick');", webObject[0]);
                    strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        //Purpose: click on an object by using CSS property to locate the object
        public bool ClickByCss(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.CssSelector(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on {0} successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        //Purpose: take a mouse over action
        public bool MouseOver(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Mouse over {0} successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        /// Purpose: Get the order of an object
        /// </summary>
        public bool GetOrderNo(IWebDriver webDriver, string oObjectXpath)
        {
            bool bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                IWebElement webOrderNo = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath), _clsCommon.IntTimeOut);
                _strOderNo = webOrderNo.GetAttribute("innerHTML");
                if (_strOderNo != null || _strOderNo == "")
                {
                    bResult = true;
                    var strStepDetails = "Get Order No successfully. The Order No is " + _strOderNo;
                    Console.WriteLine(strStepDetails);
                    Log.Info(strStepDetails);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        /// Purpose: Clear data of an object
        /// </summary>
        public bool ClearData(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    oWebElements[0].Clear();
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. Clear data for {0} successfully.",strObjectName);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        /// Purpose: Take a right mouse click action
        /// </summary>
        public bool RightClick(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.ContextClick(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Right click on {0} successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        /// Purpose: checking data belongs the object or not
        /// </summary>
        public bool VerifyText_Contain(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    string txtMessage = oWebElements[0].GetAttribute("innerHTML").ToLower().Trim();
                    //Print out message if both strData and strOderNo are blank or null
                    if ((strData.IsNullOrEmpty() || strData == "") && (_strOderNo.IsNullOrEmpty() || _strOderNo == ""))
                    {
                        strStepDetails.AppendFormat("Fail. There isn't any Test Data");
                    }
                    else
                    {
                        //Verify Text for searching function and strData is null or blank
                        if (strData.IsNullOrEmpty() || strData == "")
                        {
                            if (txtMessage.Contains(_strOderNo.ToLower().Trim()))
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. Data is matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat(
                                    "Fail. Data is not matched. Expected result is {0} but Actual result is {1}.",
                                    _strOderNo.ToUpper().Trim(), txtMessage.ToUpper().Trim());
                            }
                        }
                        // /Verify Text for searching function and strData isn't null or blank. It also use for verify text other cases
                        else
                        {
                            if (txtMessage.Contains(strData.ToLower().Trim()))
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. Data is matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat(
                                    "Fail. Data is not matched. Expected result is {0} but Actual result is {1}.",
                                    _strOderNo.ToUpper().Trim(), txtMessage.ToUpper().Trim());
                            }
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is not displayed", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        /// Purpose: checking data equal to the value of object or not
        /// </summary>
        public bool VerifyText(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            String useragent = (String)js.ExecuteScript("return navigator.userAgent;");
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    string txtMessage;
                    if (useragent.Contains("MSIE 8.0"))
                    {
                        txtMessage = oWebElements[0].Text.Trim(); // for IE 8    
                    }
                    else
                    {
                        txtMessage = oWebElements[0].GetAttribute("textContent").Trim(); // for IE 9 or above and Chrome.    
                    }
                    if (strData.IsNullOrEmpty() && txtMessage.IsNullOrEmpty())
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. The data are matched.");
                    }
                    else
                    {
                        if (txtMessage != "")
                        {
                            if (txtMessage.Equals(strData.Trim()))
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. Data are not matched. Expected result:[{0}] Actual result: [{1}]", strData.Trim(), txtMessage.Trim());
                            }
                        }
                        else
                        {
                            var value = oWebElements[0].GetAttribute("value");
                            if (value != "0" && !value.IsNullOrEmpty())
                            {
                                if (value.Equals(strData))
                                {
                                    bResult = true;
                                    strStepDetails.AppendFormat("Passed. The data are matched.");
                                }
                                else
                                {
                                    strStepDetails.AppendFormat("Fail. Data are not matched. Expected result:[{0}] Actual result: [{1}]", strData.Trim(), txtMessage.Trim());
                                }
                            }
                            else // Check
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is not displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        public bool VerifyText_Contain2(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            String useragent = (String)js.ExecuteScript("return navigator.userAgent;");
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    string txtMessage;
                    if (useragent.Contains("MSIE 8.0"))
                    {
                        txtMessage = oWebElements[0].Text.Trim(); // for IE 8    
                    }
                    else
                    {
                        txtMessage = oWebElements[0].GetAttribute("textContent").Trim(); // for IE 9 or above and Chrome.    
                    }
                    if (strData.IsNullOrEmpty() && txtMessage.IsNullOrEmpty())
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. The data are matched.");
                    }
                    else
                    {
                        if (txtMessage != "")
                        {
                            if (txtMessage.Contains(strData.Trim()))
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. Data are not matched. Expected result:[{0}] Actual result: [{1}]", strData.Trim(), txtMessage.Trim());
                            }
                        }
                        else
                        {
                            var value = oWebElements[0].GetAttribute("value");
                            if (value != "0" && !value.IsNullOrEmpty())
                            {
                                if (value.Contains(strData))
                                {
                                    bResult = true;
                                    strStepDetails.AppendFormat("Passed. The data are matched.");
                                }
                                else
                                {
                                    strStepDetails.AppendFormat("Fail. Data are not matched. Expected result:[{0}] Actual result: [{1}]", strData.Trim(), txtMessage.Trim());
                                }
                            }
                            else // Check
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is not displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }

        /// <summary>
        /// Purpose: wait for a time out
        /// </summary>
        /// <returns></returns>
        public bool WaitTime()
        {
            bool bResult = true;
            try
            {
                Thread.Sleep(_clsCommon.IntTimeOut);
                Console.WriteLine("Waitting Time: " + _clsCommon.IntTimeOut);
            }
            catch (Exception ex)
            {
                bResult = false;
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }

        /// <summary>
        /// Purpose: wait for a duration
        /// </summary>
        /// <returns></returns>
        public bool WaitSpecificTime(string strData)
        {
            bool bResult = true;
            try
            {
                int waitTime = Convert.ToInt16(strData);//in second
                Console.WriteLine("Waitting Time: " + waitTime);
                Thread.Sleep(TimeSpan.FromSeconds(waitTime));
            }
            catch (Exception ex)
            {
                bResult = false;
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        /// <summary>
        /// This function will wait an object present until timeup.
        /// </summary>
        public bool WaitAnElementPresent(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            for (int i = 0; i < _clsCommon.IntTimeSearch; i += SleepTime)
            {
                try
                {
                    ReadOnlyCollection<IWebElement> oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    if (oWebElements.Any())
                    {
                        Console.WriteLine("The element {1} is presented at {0} milliseconds.", i + SleepTime, oObjectXpath);
                        Thread.Sleep(SleepTime);
                        return true;
                    }
                    Thread.Sleep(SleepTime);
                    Console.WriteLine("Waiting for the element {1} displays at {0} milliseconds.", i + SleepTime, oObjectXpath);
                }
                catch (Exception ex)
                {
                    Log.Error(ex.ToString());
                }
            }
            Console.WriteLine("The element {1} is still not displayed after waited for {0} milliseconds.", _clsCommon.IntTimeOut, oObjectXpath);
            return false;
        }
        /// <summary>
        /// This function will wait an object disappear until timeup.
        /// </summary>
        public bool WaitAnElementNotPresent(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            for (int i = 0; i < _clsCommon.IntTimeSearch; i += SleepTime)
            {
                try
                {
                    ReadOnlyCollection<IWebElement> oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    if (oWebElements.Any())
                    {
                        Thread.Sleep(SleepTime);
                        Console.WriteLine("The element {1} is still presented at {0} milliseconds.", i + SleepTime, oObjectXpath);
                    }
                    else
                    {
                        Console.WriteLine("The element {1} is not presented at {0} milliseconds.", i + SleepTime, oObjectXpath);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Log.Error(ex.ToString());
                }
            }
            Console.WriteLine("The element {1} is presented at {0} milliseconds.", _clsCommon.IntTimeOut, oObjectXpath);
            return false;
        }
        /// <summary>
        /// This function helps to get the value
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        public bool GetValue(IWebDriver webDriver, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    _strOderNo = oWebElements[0].GetAttribute("innerHTML");
                    if (_strOderNo != null || _strOderNo == "")
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed.The Order No is {0}" ,_strOderNo);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Miss the order.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        /// This function helps to double click on the element
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="strObjectName"></param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        public bool DoubleClicks(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.DoubleClick(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Double clicks on {0} successfully.",strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is not displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        /// This function helps to scroll down the specific form
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if scroll down successfully
        ///     False: if scroll down unsuccessfully
        /// </returns>
        public bool ScrollDownForm(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.PageDown);
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. Scroll down {0} successfully.", strObjectName);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        /// This function helps to scroll up the specific form
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if scroll up successfully
        ///     False: if scroll up unsuccessfully
        /// </returns>
        public bool ScrollUpForm(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    oWebElements[0].Click();
                    oWebElements[0].SendKeys(Keys.PageUp);
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. Scroll up {0} successfully.", strObjectName);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        /// This function helps to send the Enter key to the specific object in current page
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <returns>
        ///     True: if the Enter key has been sent successfully
        ///     False: if the Enter key has not been sent unsuccessfully
        /// </returns>
        public bool PressEnter(IWebDriver webDriver)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                Actions builder = new Actions(webDriver);
                builder.SendKeys(Keys.Enter).Perform();
                strStepDetails.AppendFormat("Passed. Press Enter key successfully.");
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        public bool PressEnterOnTheObject(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.Enter);
                    strStepDetails.AppendFormat("Passed. Press Enter key successfully.");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        public bool PressEnterOnTheObject2(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.Enter);
                    strStepDetails.AppendFormat("Passed. Press Enter key successfully.");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        public bool DismissAlert(IWebDriver webDriver)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                IAlert alt = webDriver.SwitchTo().Alert();
                alt.Dismiss();
                strStepDetails.AppendFormat("Passed. Dismiss the Alert successfully.");
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        public bool AcceptAlert(IWebDriver webDriver)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                IAlert alt = webDriver.SwitchTo().Alert();
                alt.Accept();
                strStepDetails.AppendFormat("Passed. Accept the Alert successfully.");
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        /// This function helps to send the Tab key to the specific object in current page
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <returns>
        ///     True: if the Tab key has been sent successfully
        ///     False: if the tab key has not been sent unsuccessfully
        /// </returns>
        public bool PressTab(IWebDriver webDriver)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                Actions builder = new Actions(webDriver);
                builder.SendKeys(Keys.Tab).Perform();
                strStepDetails.AppendFormat("Passed. Press Tab key successfully.");
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        public bool PressTabOnTheObject(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.Tab);
                    strStepDetails.AppendFormat("Passed. Press Tab key successfully.");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// <summary>
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Sep-2015
        /// This function helps to send the F5 key to the specific object in current page
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if the Tab key has been sent successfully
        ///     False: if the tab key has not been sent unsuccessfully
        /// </returns>

        public bool PressF5(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.F5);
                    webDriver.Navigate().Refresh();
                    strStepDetails.AppendFormat("Passed. Press F5 key successfully.");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        /// This function helps to send the Tab key to the specific object in current page
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if the ArrowDown key has been sent successfully
        ///     False: if the ArrowDown key has not been sent unsuccessfully
        /// </returns>

        public bool PressArrowDown(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    oWebElements[0].SendKeys(Keys.ArrowDown);
                    strStepDetails.AppendFormat("Passed. Press ArrowDown key successfully.");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        /// This function helps to check the status of expected object is disabled.
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if the status of object is disabled
        ///     False: if the status of object is enabled
        /// </returns>
        public bool IsButtonDisabled(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    if (!oWebElements[0].Enabled)
                    {
                        strStepDetails.AppendFormat("Passed. The {0} is disable.",strObjectName);
                        bResult = true;
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. The {0} is enable.", strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. {0} is NOT displayed.", strObjectName);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return bResult;
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// <summary>
        /// This function helps to check the status of expected object is enabled.
        /// </summary>
        /// <param name="webDriver">Selenium Web Driver</param>
        /// <param name="oObjectXpath">The xpath of specific object which was defined in or.properties</param>
        /// <param name="strObjectName">Name of object</param>
        /// <returns>
        ///     True: if the status of object is enabled.
        ///     False: if the status of object is disabled.
        /// </returns>
        public bool IsButtonEnabled(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    bool status = oWebElements[0].Enabled;
                    if (status)
                    {
                        strStepDetails.AppendFormat("Passed. The {0} is enable.", strObjectName);
                        bResult = true;
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. The {0} is disable.", strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
    //Purpose: Check an element is readonly or not
    //Written: Viet.LeMinh2@harveynash.vn
    //Date: June-2015
    public bool IsElementReadonly(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    if (oWebElements[0].GetAttribute("readonly") == "true" || oWebElements[0].Enabled == false)
                    {
                        strStepDetails.AppendFormat("Passed. The {0} is read only", strObjectName);
                        bResult = true;
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. The {0} is NOT read only.", strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
    }
    }
