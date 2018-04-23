using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Castle.Core.Internal;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Selenium_Automated_Testing.Utilities;

namespace Worksmart_Automated_Testing.Utilities
{
    public partial class Keyword
    {
        private readonly string _conn = Common.ConnectionString;
        public Dictionary<string, string> OldData = new Dictionary<string, string>();
        public Dictionary<string, string> NewData = new Dictionary<string, string>();
        const int SleepTime = 1000;
        /// summary
        ///DeleteRule() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2016
        ///Purpose: 
        /// Delete a rule of Auto Change table 
        ///Description: 
        ///input : xPath 
        ///output: Delete the rule if it exists
        /// summary
        public bool DeleteRule(IWebDriver webDriver)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            const string msgNorulesfound = "//*[@id='GridTable']//td[contains(text(),'No rules found')]";
            const string icoDelete = "//*[@id='GridTable']//i[@title='Click to Delete']";
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> oWebMsgElements = webDriver.FindElements(By.XPath(msgNorulesfound));
                ReadOnlyCollection<IWebElement> oWebDelElements = webDriver.FindElements(By.XPath(icoDelete));
                while (!oWebMsgElements.Any() && oWebDelElements.Any()) 
                    {
                        Actions builder = new Actions(webDriver);
                        builder.MoveToElement(oWebDelElements[0]).Build().Perform();
                        builder.Click(oWebDelElements[0]).Build().Perform();
                        Thread.Sleep(SleepTime);
                        oWebMsgElements = webDriver.FindElements(By.XPath(msgNorulesfound));    
                        oWebDelElements = webDriver.FindElements(By.XPath(icoDelete));
                    }
                strStepDetails.AppendFormat("Passed. The rule is deleted.");
                bResult = true;
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


        /// summary
        ///UnSelectCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Sep-2015
        ///Purpose: 
        /// Click on a check box if it is checked, do nothing if check box has been checked before 
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool UnSelectCheckBox(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
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
                    if (oWebElements[0].Selected)
                    {
                        Actions builder = new Actions(webDriver);
                        builder.MoveToElement(oWebElements[0]).Build().Perform();
                        builder.Click(oWebElements[0]).Build().Perform();
                    }
                    strStepDetails.AppendFormat("Passed. The {0} is unchecked.", strObjectName);
                    bResult = true;
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

        /// summary
        ///SelectCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Sep-2015
        ///Purpose: 
        /// Click on a check box if it is not checked, do nothing if check box has been checked before 
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool SelectCheckBox(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
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
                    if (!oWebElements[0].Selected)
                    {
                        Actions builder = new Actions(webDriver);
                        builder.MoveToElement(oWebElements[0]).Build().Perform();
                        builder.Click(oWebElements[0]).Build().Perform();
                    }
                    strStepDetails.AppendFormat("Passed. The {0} is checked.", strObjectName);
                    bResult = true;
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
        /// summary
        ///ClickOnItemOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Mar-2016
        ///Purpose: 
        /// Click on an item of "Handler Type(s) can own case at this Milestone" field of "Milestones" page in "Case Management"
        /// Click on an item of "Relevant to Process Step(s)" field of "Case Type" page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfCheckBoxAccuracy(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//label[text()='{0}']/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (!oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            //action.Click(oWebElements[0]).Build().Perform();
                            //action.SendKeys(Keys.Enter);
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is checked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///ClickOnItemOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Sep-2015
        ///Purpose: 
        /// Click on an item of "Handler Type(s) can own case at this Milestone" field of "Milestones" page in "Case Management"
        /// Click on an item of "Relevant to Process Step(s)" field of "Case Type" page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//label[contains(.,'{0}')]/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (!oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            //action.Click(oWebElements[0]).Build().Perform();
                            //action.SendKeys(Keys.Enter);
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is checked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///UnClickOnItemOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Oct-2015
        ///Purpose: 
        /// unClick on an item of "Handler Type(s) can own case at this Milestone" field of "Milestones" page in "Case Management"
        /// unClick on an item of "Relevant to Process Step(s)" field of "Case Type" page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool UnClickOnItemOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//label[contains(.,'{0}')]/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    if (oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        action.MoveToElement(oWebElements[0]).Build().Perform();
                        action.Click(oWebElements[0]).Build().Perform(); 
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. unClick on [{0}] successfully.", strData);
                    }
                    else
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. The [{0}] is unchecked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfCheckBox2() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2016
        ///Purpose: 
        /// Click on an item of "Handler Type(s) can own case at this Milestone" field of "Milestones" page in "Case Management"
        /// Click on an item of "Relevant to Process Step(s)" field of "Case Type" page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfCheckBox2(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
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
                    bResult = true;
                    if (!oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                            //action.Click(oWebElements[0]).Build().Perform();
                            //action.SendKeys(Keys.Enter);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strObjectName);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is checked.", strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strObjectName);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///UnClickOnItemOfCheckBox2() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Oct-2015
        ///Purpose: 
        /// unClick on an item of "Handler Type(s) can own case at this Milestone" field of "Milestones" page in "Case Management"
        /// unClick on an item of "Relevant to Process Step(s)" field of "Case Type" page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool UnClickOnItemOfCheckBox2(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
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
                    if (oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        action.MoveToElement(oWebElements[0]).Build().Perform();
                        action.Click(oWebElements[0]).Build().Perform();
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. unClick on [{0}] successfully.", strObjectName);
                    }
                    else
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. The [{0}] is unchecked.", strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strObjectName);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfSearchFieldOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2015
        ///Purpose: 
        /// Click on an item of "Search" field of "Customer Search" admin page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if un-click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfSearchFieldOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='customerSearchSettingForm']/div[2]//div[contains(.,'{0}')]/label/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (!oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is checked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///UnClickOnItemOfSearchFieldOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2015
        ///Purpose: 
        /// Click on an item of "Search" field of "Customer Search" admin page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if un-click is successfull and vice verse
        /// summary
        public bool UnClickOnItemOfSearchFieldOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='customerSearchSettingForm']/div[2]//div[contains(.,'{0}')]/label/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Unclicking on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is unchecked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfSearchResultOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2015
        ///Purpose: 
        /// Click on an item of "Search Result" field of "Customer Search" admin page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if un-click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfSearchResultOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='customerSearchSettingForm']/div[3]//div[contains(.,'{0}')]/label/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (!oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is checked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///UnClickOnItemOfSearchResultOfCheckBox() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Dec-2015
        ///Purpose: 
        /// Click on an item of "Search Result" field of "Customer Search" admin page in "Case Management"
        ///Description: 
        ///input : name value
        ///output: is true if un-click is successfull and vice verse
        /// summary
        public bool UnClickOnItemOfSearchResultOfCheckBox(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='customerSearchSettingForm']/div[3]//div[contains(.,'{0}')]/label/input", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    if (oWebElements[0].Selected)
                    {
                        Actions action = new Actions(webDriver);
                        if (js != null)
                        {
                            js.ExecuteScript("scroll(0, -250);");
                            action.MoveToElement(oWebElements[0]).Build().Perform();
                            js.ExecuteScript("arguments[0].click()", oWebElements[0]);
                        }
                        strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Passed. The [{0}] is unchecked.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        public bool GetWid(IWebDriver webDriver)
        {
            try
            {
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return true;
        }
        /// <summary>
        ///HoverMouse() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: August-2015
        ///Purpose: Mouse hover on a link
        /// </summary>
        public bool HoverMouse(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            bool bResult = false;
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
                    strStepDetails.AppendFormat("Passed. Hover mouse on [{0}] successfully.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is NOT displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
            return bResult;
        }
        /// <summary>
        ///ClickOnLastPage() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: August-2015
        ///Purpose: Click on the last page of Case Outcomes, Customer Outcomes, Handler Types, Letter Template...
        /// In Parameter: xPath, table name
        /// Output: Click on the last page successfully or not
        /// </summary>

        public bool ClickOnLastPage(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            var oObjectXpath = new StringBuilder();
            var sqlCmd = new StringBuilder();
            var nameArr = strData.Split('|');
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                using (var connection = new SqlConnection(_conn))
                {
                    using (var cmd = connection.CreateCommand())
                    {
                        if (nameArr.Length.Equals(1))
                        {
                            sqlCmd.Clear();
                            sqlCmd.AppendFormat("SELECT CASE COUNT(*)%20 WHEN 0 THEN COUNT(*)/20 ELSE ROUND(COUNT(*)/20,0) + 1 END  FROM {0}", nameArr[0]);
                        }
                        else
                        {
                            sqlCmd.Clear();
                            sqlCmd.AppendFormat("SELECT CASE COUNT(*)%20 WHEN 0 THEN COUNT(*)/20 ELSE ROUND(COUNT(*)/20,0) + 1 END  FROM {0} WHERE CaseTypeID = (SELECT CaseTypeId FROM dbo.CM_CaseTypeList WHERE CaseTypeCode='{1}')", nameArr[0], nameArr[1]);
                        }
                        cmd.CommandText = sqlCmd.ToString();
                        cmd.CommandType = CommandType.Text;
                        connection.Open();
                        var countResult = (int)cmd.ExecuteScalar();
                        connection.Close();
                        if (countResult >= 2)
                        {
                            oObjectXpath.AppendFormat("//*[@class='pagination']//a[text()='{0}']", countResult);
                            int i = 0;
                            ReadOnlyCollection<IWebElement> oWebElements;
                            do
                            {
                                oWebElements = webDriver.FindElements(By.XPath(oObjectXpath.ToString()));
                                i += SleepTime;
                                Thread.Sleep(SleepTime);
                            } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                            if (oWebElements.Any())
                            {
                                Actions builder = new Actions(webDriver);
                                builder.MoveToElement(oWebElements[0]).Build().Perform();
                                builder.Click(oWebElements[0]).Build().Perform();
                                strStepDetails.AppendFormat("Passed. Click on the last page successfully.");
                                bResult = true;
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. The last page is NOT displayed.");
                            }
                        }
                    }
                }
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
        ///VerifyPageTitle() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: August-2015
        ///Purpose: return the row number of table divide 20
        /// In Parameter: xPath, table name
        /// Output: number of row
        /// </summary>

        public bool VerifyPageTitle(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            var sqlCmd = new StringBuilder();
            var nameArr = strData.Split('|');
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                using (var connection = new SqlConnection(_conn))
                {
                    using (var cmd = connection.CreateCommand())
                    {
                        if (nameArr.Length.Equals(1))
                        {
                            sqlCmd.Clear();
                            sqlCmd.AppendFormat("SELECT CASE COUNT(*)%20 WHEN 0 THEN COUNT(*)/20 ELSE ROUND(COUNT(*)/20,0) + 1 END  FROM {0}", nameArr[0]);
                        }
                        else
                        {
                            sqlCmd.Clear();
                            sqlCmd.AppendFormat("SELECT CASE COUNT(*)%20 WHEN 0 THEN COUNT(*)/20 ELSE ROUND(COUNT(*)/20,0) + 1 END  FROM {0} WHERE CaseTypeID = (SELECT CaseTypeId FROM dbo.CM_CaseTypeList WHERE CaseTypeCode='{1}')", nameArr[0], nameArr[1]);
                        }
                        cmd.CommandText = sqlCmd.ToString();
                        cmd.CommandType = CommandType.Text;
                        connection.Open();
                        var countResult = (int)cmd.ExecuteScalar();
                        connection.Close();
                        if (countResult >= 2)
                        {
                            int i = 0;
                            ReadOnlyCollection<IWebElement> oWebElements;
                            do
                            {
                                oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                                i += SleepTime;
                                Thread.Sleep(SleepTime);
                            } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                            if (oWebElements.Any())
                            {
                                string txtMessage = oWebElements[0].GetAttribute("innerHTML").Trim();
                                if (txtMessage.Contains(countResult.ToString()))
                                {
                                    bResult = true;
                                    strStepDetails.AppendFormat("Passed. The title of paging is matched.");
                                }
                                else
                                {
                                    strStepDetails.AppendFormat(
                                        "Fail. Data is not matched. Expected result is {0} but Actual result is {1}.",
                                        countResult.ToString().ToUpper().Trim(), txtMessage.ToUpper().Trim());
                                }
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                            }
                        }
                    }
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
        ///VerifyContentOfBreadCrumb() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: return content of before part of an element
        ///Description: use JQuery command and execute by JavaScript command
        /// </summary>
        public bool VerifyContentOfBreadCrumb(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    var content = GetContentFromElement(webDriver, oObjectXpath, false);
                    if (content.Contains(strData))
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. [{0}] is displayed but The title is not right.", strData);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
            return bResult;
        }
        public string GetContentFromElement(IWebDriver webDriver, string oObjectXpath, bool isAfter)
        {
            var js = webDriver as IJavaScriptExecutor;
            string scriptContent;
            Thread.Sleep(_clsCommon.IntTimeWait);
            if (isAfter)
            {
                scriptContent = "var element = document.evaluate(\"" + oObjectXpath +
                                "\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
                                "return window.getComputedStyle(element,':after').getPropertyValue('content')";
            }
            else
            {
                //for IE 9, 10,11, Chrome
                scriptContent = "var element = document.evaluate(\"" + oObjectXpath +
                                "\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
                                "return window.getComputedStyle(element,':before').getPropertyValue('content')";
            }
            if (js != null)
            {
                return (String)js.ExecuteScript(scriptContent);
            }
            return String.Empty;

        }

        /// summary
        ///VerifyItemOfActionOutcomesIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Action Outcomes" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfActionOutcomesIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='actionOutcomeTable']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strData);
                }
                else
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///VerifyItemOfActionOutcomesIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Action Outcome" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfActionOutcomesIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='actionOutcomeTable']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfActionOutcomes() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Action Outcomes" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfActionOutcomes(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='actionOutcomeTable']//a[contains(text(),'{0}')]", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                    IWebElement webInput = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath),
                        _clsCommon.IntTimeOut);
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(webInput).Build().Perform();
                    webInput.Click();
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }

                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///VerifyItemOfPrioritiesIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "priorities" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfPrioritiesIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='priorityTable']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strData);
                }
                else
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///VerifyItemOfPrioritiesIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Priorities" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfPrioritiesIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='priorityTable']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnPageOfLetterTemplate() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Click on an a page of navigation of the "Letter Template" list page
        ///Description: 
        ///input : namuber of page
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnPageOfLetterTemplate(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='gridLetterTemplate']//a[text()='{0}')]", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. Page [{0}] is clicked.", strData);
                    IWebElement webInput = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath),
                        _clsCommon.IntTimeOut);
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(webInput).Build().Perform();
                    webInput.Click();
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Page [{0}] is not clicked.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///ClickOnPageOfNoneHierarchicalList() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Click on an a page of navigation of the none hierarchical list page
        ///Description: 
        ///input : number of page
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnPageOfNoneHierarchicalList(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                //var oObjectXpath = string.Format("//*[@id='plData']//a[contains(text(),'{0}')]", strData);
                var oObjectXpath = string.Format("//*[@class='pagination']//a[contains(text(),'{0}')]", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. Page [{0}] is clicked.", strData);
                    IWebElement webInput = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath),
                        _clsCommon.IntTimeOut);
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(webInput).Build().Perform();
                    webInput.Click();
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Page [{0}] is NOT displayed.", strData);
                }

                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfPriorities() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Priorities" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfPriorities(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='priorityTable']//a[contains(text(),'{0}')]", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                    IWebElement webInput = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath),
                        _clsCommon.IntTimeOut);
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(webInput).Build().Perform();
                    webInput.Click();
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// summary
        ///VerifyItemOfCaseStatesIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Case States" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfCaseStatesIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strData);
                }
                else
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is NOT displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// summary
        ///VerifyItemOfNonHierarchicalListIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Aug-2015
        ///Purpose: Click on an item of table of "Case States" or "Priorities" or "Action Outcomes" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfNonHierarchicalListIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strData);
                }
                else
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is NOT displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// summary
        ///VerifyItemOfCaseStatesIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Case States" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfCaseStatesIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is not displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// summary
        ///VerifyItemOfNonHierarchicalListIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Aug-2015
        ///Purpose: Click on an item of table of "Case States" or "Priorities" or "Action Outcomes" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyItemOfNonHierarchicalListIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[text()='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is not displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// summary
        ///ClickOnItemOfNonHierarchicalList() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Aug-2015
        ///Purpose: Click on an item of table of "Case States"  or "Priorities" or "Action Outcomes" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfNonHierarchicalList(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[contains(.,'{0}')]", strData);
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
                    //action.Click(oWebElements[0]).Build().Perform();
                    action.DoubleClick(oWebElements[0]).Build().Perform();
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is NOT displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }

        /// summary
        ///ClickOnItemOfCaseStates() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on an item of table of "Case States" page
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfCaseStates(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@id='pageContent']//a[contains(text(),'{0}')]", strData);
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
                    action.Click(oWebElements[0]).Build().Perform();
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is not displayed.", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// summary
        ///ClickOnCheckOfAvailableNature() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on one or more item of tree in "Available Nature" 
        ///Description: 
        ///input : name value list, '|' mark is used to seperate values
        ///output: selected list of item
        /// summary
        public bool ClickOnCheckOfAvailableNature(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var nameArr = strData.Split('|');
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                const string xPathRoot = "//*[@id='plNatureTree']//a[text()='►']";
                ReadOnlyCollection<IWebElement> oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                while (oWebArrowElements.Any())
                {
                    _clsCommon.FindElement(webDriver, By.XPath(xPathRoot), _clsCommon.IntTimeOut).Click();
                    int i = 0;
                    do
                    {
                        oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                        i += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebArrowElements.Any() && i < _clsCommon.IntTimeSearch);

                }
                foreach (string t in nameArr)
                {
                    var oObjectXpath = "//*[@id='plNatureTree']//input[@text='" + t + "']";
                    IWebElement oClick = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath), _clsCommon.IntTimeOut);
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oClick).Build().Perform();
                    builder.Click(oClick).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on [{0}] is successful.", strData);
                    Thread.Sleep(_clsCommon.IntClickWait); 
                    Console.WriteLine(strStepDetails);
                    Log.Info(strStepDetails);
                    bResult = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///VerifyItemOftreeIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Click on 'Add Child Item' of item of tree (e.g. worklist, product, nature, shared lookup type)
        ///Description: 
        ///input : tree value
        ///output: is true if item is absent and vice verse
        /// summary
        public bool VerifyItemOftreeIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = "//*[@id='tree2']//span[text()='" + strData + "']";
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                const string xPathRoot = "//*[@id='tree2']//a[text()='►']";
                ReadOnlyCollection<IWebElement> oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                while (!oWebElements.Any() && oWebArrowElements.Any())
                {
                    oWebArrowElements[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);
                    oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                }
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. The {0} is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Passed. The {0} is not displayed.", strData);
                    bResult = true;
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
        /// summary
        ///VerifyItemOftreeIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Click on 'Add Child Item' of item of tree (e.g. worklist, product, nature, shared lookup type)
        ///Description: 
        ///input : tree value
        ///out put: is true if item is present and vice verse
        /// summary
        public bool VerifyItemOftreeIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = "//*[@id='tree2']//span[text()='" + strData + "']";
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                const string xPathRoot = "//*[@id='tree2']//a[text()='►']";
                ReadOnlyCollection<IWebElement> oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                while (!oWebElements.Any() && oWebArrowElements.Any())
                {
                    oWebArrowElements[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);
                    oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                }
                if (oWebElements.Any())
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. The {0} is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
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

        /// summary
        ///ClickOnAddChildItem() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Click on 'Add Child Item' of item of tree (e.g. worklist, product, nature, shared lookup type)
        ///Description: 
        ///input : tree value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnAddChildItem(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                var oObjectXpath = "//*[@id='tree2']//span[text()='" + strData + "']/following-sibling::*[a]";
                int j = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    j += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);

                const string xPathRoot = "//*[@id='tree2']//a[text()='►']";
                ReadOnlyCollection<IWebElement> oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                while (!oWebElements.Any() && oWebArrowElements.Any())
                {
                    oWebArrowElements[0].Click();
                    int i = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        i += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                    oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                }
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    action.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on [{0}] successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Click on [{0}] unsuccessfully.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait);
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
        /// summary
        ///ClickOnNameOfItemInLetterTemplateTable() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Click on 'Name' of item of table (none hierarchical shared lookup type)
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnNameOfItemInLetterTemplateTable(IWebDriver webDriver, string strData)
        {
            var strStepDetails = new StringBuilder();
            bool bResult = false;
            int pageNumber = 1;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var nameXpath = string.Format("//*[@id='listLetterTemplate']//a[contains(text(),'{0}')]", strData);
                int j = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(nameXpath));
                    j += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);
                var pageXpath = string.Format("//*[@id='listLetterTemplate']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oWebNavigationElements = webDriver.FindElements(By.XPath(pageXpath));
                while (oWebNavigationElements.Any() && !oWebElements.Any())
                {
                    oWebNavigationElements[0].Click();
                    int i = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(pageXpath));
                        i += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                    pageNumber = pageNumber + 1;
                    pageXpath = string.Format("//*[@id='listLetterTemplate']//a[text()='{0}']", pageNumber);
                    oWebNavigationElements = webDriver.FindElements(By.XPath(pageXpath));
                }
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Passed. Click on the {0} successfully.", strData);
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    action.Click(oWebElements[0]).Build().Perform();
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            return bResult;
        }
        /// summary
        ///ClickOnCodeOfSelectActionDefinitionDialog() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on 'Name' of code column of table (Select Action Definition dialog)
        ///Description: 
        ///input : code value
        ///output: is true if click is successfull and vice verse
        /// summary
        /// 
        public bool ClickOnCodeOfSelectActionDefinitionDialog(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            var xpathPage = new StringBuilder();
            int pageNumber = 1;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered ellipsis-table']//span/a[contains(text(),'{0}')]", strData);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeWait);
                xpathPage.Clear();
                xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                while (oObjectPage.Any() && !oWebElements.Any())
                {
                    oObjectPage[0].Click();
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeWait);
                    pageNumber = pageNumber + 1;
                    xpathPage.Clear();
                    xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                    oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                }
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    action.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on the [{0}] successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///ClickOnNameOfItem() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on 'Name' of item of table (none hierarchical shared lookup type)
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        /// 
        public bool ClickOnNameOfItem(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            var xpathPage = new StringBuilder();
            int pageNumber = 1;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                //var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[contains(text(),'{0}')]", strData);
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[text()='{0}']", strData);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeWait);
                xpathPage.Clear();
                xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                while (oObjectPage.Any() && !oWebElements.Any())
                {
                    oObjectPage[0].Click();
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeWait);
                    pageNumber = pageNumber + 1;
                    xpathPage.Clear();
                    xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                    oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                }
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    action.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on the [{0}] successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait);
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnNameOfItem2() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on 'Name' of item of table (none hierarchical shared lookup type) with space 
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        /// 
        public bool ClickOnNameOfItem2(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            var xpathPage = new StringBuilder();
            int pageNumber = 1;
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[contains(text(),'{0}')]", strData);
                //var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[text()='{0}']", strData);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeWait);
                xpathPage.Clear();
                xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                while (oObjectPage.Any() && !oWebElements.Any())
                {
                    oObjectPage[0].Click();
                    if (js != null) js.ExecuteScript("scroll(0, -250);");
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeWait);
                    pageNumber = pageNumber + 1;
                    xpathPage.Clear();
                    xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                    oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                }
                if (oWebElements.Any())
                {
                    Actions action = new Actions(webDriver);
                    action.MoveToElement(oWebElements[0]).Build().Perform();
                    action.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on the [{0}] successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait);
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnNameOfItemOnCurrentPage() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Sep-2015
        ///Purpose: Click on 'Name' of item of table (none hierarchical shared lookup type)
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        /// 
        public bool ClickOnNameOfItemOnCurrentPage(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[contains(text(),'{0}')]", strData);
                int i = 0;
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
                    action.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on the [{0}] successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// summary
        ///VerifyNameIsDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on 'Name' of item of table (shared lookup type)
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyNameIsDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            var xpathPage = new StringBuilder();
            int pageNumber = 1;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[text()='{0}']", strData);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                xpathPage.Clear();
                xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                while (oObjectPage.Any() && !oWebElements.Any())
                {
                    oObjectPage[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);

                    pageNumber++;
                    xpathPage.Clear();
                    xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                    oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                }
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Passed. The [{0}] is displayed.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is NOT displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///VerifyNameIsNotDisplayed() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on 'Name' of item of table (shared lookup type)
        ///Description: 
        ///input : name value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool VerifyNameIsNotDisplayed(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            var xpathPage = new StringBuilder();
            int pageNumber = 1;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = string.Format("//*[@class='table admin-table table-striped table-bordered']//a[text()='{0}']", strData);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                xpathPage.Clear();
                xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                ReadOnlyCollection<IWebElement> oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                while (oObjectPage.Any() && !oWebElements.Any())
                {
                    oObjectPage[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);

                    pageNumber++;
                    xpathPage.Clear();
                    xpathPage.AppendFormat("//*[@class='pagination']//a[text()='{0}']", pageNumber);
                    oObjectPage = webDriver.FindElements(By.XPath(xpathPage.ToString()));
                }
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. The [{0}] is displayed.", strData);
                }
                else
                {
                    strStepDetails.AppendFormat("Passed. [{0}] is not displayed.", strData);
                    bResult = true;
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemOfTree() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Click on item of tree (e.g. worklist, product, nature, shared lookup type)
        ///Description: 
        ///input : tree value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemOfTree(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var oObjectXpath = "//*[@id='tree2']//span[text()='" + strData + "']";
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeWait);
                const string xPathRoot = "//*[@id='tree2']//a[text()='►']";
                ReadOnlyCollection<IWebElement> oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                while (!oWebElements.Any() && oWebArrowElements.Any())
                {
                    oWebArrowElements[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeWait);
                    oWebArrowElements = webDriver.FindElements(By.XPath(xPathRoot));
                }
                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on {0} successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed", strData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            Thread.Sleep(_clsCommon.IntClickWait); 
            Console.WriteLine(strStepDetails);
            Log.Info(strStepDetails);
            return bResult;
        }
        /// summary
        ///ClickOnItemForQuickSearch() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: June-2015
        ///Purpose: Click on item for quick search (in product, Action Definitions-> Advanced)
        ///Description: 
        ///input : name of tree value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemForQuickSearch(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var namexPath = "//*[@id='page-content-wrapper']//span[text()='" + strData + "']";
                const string arrowxPath = "//*[@id='page-content-wrapper']//a[text()='►']";
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(namexPath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                ReadOnlyCollection<IWebElement> oWebArrows = webDriver.FindElements(By.XPath(arrowxPath));
                while (!oWebElements.Any() && oWebArrows.Any())
                {
                    oWebArrows[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(namexPath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);
                    oWebArrows = webDriver.FindElements(By.XPath(arrowxPath));
                }
                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on {0} successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }
        /// summary
        ///ClickOnItemForQuickSearchOfWorklist() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Oct-2015
        ///Purpose: Click on item for quick search of "Create New Case" page
        ///Description: 
        ///input : name of tree value
        ///output: is true if click is successfull and vice verse
        /// summary
        public bool ClickOnItemForQuickSearchOfWorklist(IWebDriver webDriver, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                var namexPath = "//*[@id='treeDiv']//span[text()='" + strData + "']";
                const string arrowxPath = "//*[@id='treeDiv']//a[text()='►']";
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(namexPath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                ReadOnlyCollection<IWebElement> oWebArrows = webDriver.FindElements(By.XPath(arrowxPath));
                while (!oWebElements.Any() && oWebArrows.Any())
                {
                    oWebArrows[0].Click();
                    int j = 0;
                    do
                    {
                        oWebElements = webDriver.FindElements(By.XPath(namexPath));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebElements.Any() && j < _clsCommon.IntTimeSearch);
                    oWebArrows = webDriver.FindElements(By.XPath(arrowxPath));
                }
                if (oWebElements.Any())
                {
                    Actions builder = new Actions(webDriver);
                    builder.MoveToElement(oWebElements[0]).Build().Perform();
                    builder.Click(oWebElements[0]).Build().Perform();
                    strStepDetails.AppendFormat("Passed. Click on {0} successfully.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strData);
                }
                Thread.Sleep(_clsCommon.IntClickWait); 
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Got an exception! " + ex);
                Log.Error(ex.ToString());
                bResult = false;
            }
            return bResult;
        }

        /// Purpose: Verify file name  after uploading
        public bool VerifyUploadFileName(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                ReadOnlyCollection<IWebElement> element = webDriver.FindElements(By.XPath(oObjectXpath));
                if (element.Any())
                {
                    var txtMessage = element[0].GetAttribute("value").Trim();
                    // ReSharper disable once StringLastIndexOfIsCultureSpecific.1
                    string filename = txtMessage.Substring(txtMessage.LastIndexOf("\\") + 1);
                    txtMessage = filename;
                    if (strData.IsNullOrEmpty())
                    {
                        strStepDetails.AppendFormat("Fail. There isn't test data for file name");
                    }
                    else
                    {
                        if (txtMessage.Equals(strData.Trim()))
                        {
                            bResult = true;
                            strStepDetails.AppendFormat("Passed. File names are  matched.");
                        }
                        else
                        {
                            strStepDetails.AppendFormat("Fail. Data are not matched. Expected result is {0} but Actual result is {1}", strData.Trim(), txtMessage.Trim());
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("{0} is not displayed.", strObjectName);
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

        // purpose: checking the order of data in table
        // Written: Viet.LeMinh2@harveynash.vn
        //Date: May-2015
        private bool IsTableSorted(IList<IWebElement> elements, bool isAcsending)
        {
            for (var j = 0; j < elements.Count - 1; j++)
            {
                var result = String.Compare(elements[j].Text, elements[j + 1].Text, StringComparison.InvariantCulture);
                if (result > 0 && isAcsending || result < 0 && !isAcsending)
                {
                    return false;
                }
            }
            return true;
        }

        // purpose: checking the order of data in table
        // Written: Viet.LeMinh2@harveynash.vn
        //Date: May-2015
        public bool CheckOrderOfDataOnOnePage(IWebDriver webDriver, string strData)
        {
            try
            {
                const string xPathAscOrderIcon = "//*[@id='gridLetterTemplate']//a/i[@data-sort-direction='Ascending']";
                var xPathData = "//*[@id='gridLetterTemplate']/table/tbody/tr[..]/td[" + strData + "]";
                var elementCode = _clsCommon.FindElement(webDriver, By.XPath(xPathData), _clsCommon.IntTimeOut);
                IList<IWebElement> elemenCodetList = elementCode.FindElements(By.XPath(xPathData));

                var oWebElements = webDriver.FindElements(By.XPath(xPathAscOrderIcon));
                var isAscending = oWebElements.Any();
                var isSorted = IsTableSorted(elemenCodetList, isAscending);
                var strStepDetails = isSorted ? "Passed. The order is correct" : "Failed. The order is wrong.";
                Console.WriteLine(strStepDetails);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
        }
        // purpose: Get the sort order of a node in tree
        // Written: Viet.LeMinh2@harveynash.vn
        //Date: June-2015
        private int GetSortOrder(IWebDriver webDriver, string oObjectXpath, string listType)
        {
            int valueReturn;
            int i = 0;
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
                var webMessage = _clsCommon.FindElement(webDriver, listType == "1" ? By.XPath("//*[@id='SortOrder']") : By.XPath("//*[@id='NatureItem_SortOrder']"), _clsCommon.IntTimeOut);
                valueReturn = int.Parse(webMessage.GetAttribute("value").Trim());
                int j = 0;
                ReadOnlyCollection<IWebElement> oWebElementsCancel;
                do
                {
                    oWebElementsCancel = webDriver.FindElements(By.XPath(oObjectXpath));
                    j += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElementsCancel.Any() && j < _clsCommon.IntTimeOut);
                if (oWebElementsCancel.Any())
                {
                    oWebElementsCancel[0].Click();
                }
            }
            else
            {
                valueReturn = 0;
            }
            return valueReturn;
        }
        // purpose: checking the order of data in tree
        // Written: Viet.LeMinh2@harveynash.vn
        //Date: June-2015
        private bool IsTreeSorted(IWebDriver webDriver, bool isAcsending, string listType)
        {
            int i = 2;
            try
            {
                var element = _clsCommon.FindElement(webDriver, By.XPath("//*[@id='plData']/ul"), _clsCommon.IntTimeOut);
                IList<IWebElement> allItem = element.FindElements(By.XPath("//li[@class='jqtree_common']"));
                while (i < allItem.Count)
                {
                    var j = i - 1;
                    var xPath = "//*[@id='plData']//li[" + j + "]//span[1]";
                    var xPathNext = "//*[@id='plData']//li[" + i + "]//span[1]";
                    var item = _clsCommon.FindElement(webDriver, By.XPath(xPath), _clsCommon.IntTimeOut).Text;
                    var itemNext = _clsCommon.FindElement(webDriver, By.XPath(xPathNext), _clsCommon.IntTimeOut).Text;
                    //get sort order
                    var sortOrder = GetSortOrder(webDriver, xPath, listType);
                    //get sort order next
                    var sortOrderNext = GetSortOrder(webDriver, xPathNext, listType);
                    //compare by sort order
                    int resultSortOrder;
                    if (sortOrder < sortOrderNext)
                    {
                        resultSortOrder = -1;
                    }
                    else if (sortOrder > sortOrderNext)
                    {
                        resultSortOrder = 1;
                    }
                    else //sprtOrder = sortOrderNext
                    {
                        resultSortOrder = 0;
                    }
                    //if (resultSortOrder > 0 && isAcsending || resultSortOrder < 0 && !isAcsending)
                    if (resultSortOrder == 1 && isAcsending || resultSortOrder == -1 && !isAcsending)
                    {
                        return false;
                    }
                    // compare by name
                    if (resultSortOrder == 0)
                    {
                        var result = String.Compare(item, itemNext, StringComparison.InvariantCulture);
                        if (result > 0 && isAcsending || result < 0 && !isAcsending)
                        {
                            return false;
                        }
                    }
                    i = i + 1;
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
        }
        // purpose: checking the order of data in table
        // Written: Viet.LeMinh2@harveynash.vn
        //Date: June-2015
        //strData= 1 -> Worklist and product, shared lookup types
        //=0 -> nature
        public bool CheckOrderOfDataOnOneTree(IWebDriver webDriver, string strData)
        {
            try
            {
                const string xPathAscOrderIcon = "//*[@id='selSortDirection']";
                var webMessage = _clsCommon.FindElement(webDriver, By.XPath(xPathAscOrderIcon), _clsCommon.IntTimeOut);
                var selectedValue = new SelectElement(webMessage);
                var txtMessage = selectedValue.SelectedOption.Text;
                bool isAscending = txtMessage != "Descending";
                var isSorted = IsTreeSorted(webDriver, isAscending, strData);
                var strStepDetails = isSorted ? "Passed. The order is correct" : "Failed. The order is not correct.";
                Console.WriteLine(strStepDetails);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
        }
        public bool CheckColorOfAllPageButton(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            const string xPath = ".//*[@id='gridLetterTemplate']/nav/ul/li[..]/a[text()='";
            var bResult = false;
            var rowCount = CompareLetterTemplateTable("");
            var row = Convert.ToInt64(rowCount);
            var rowCountI = (int)row / 50;
            if (row % 50 > 0)
            {
                rowCountI += 1;
            }
            try
            {
                if (_clsCommon.FindElement(webDriver, By.XPath(oObjectXpath), _clsCommon.IntTimeOut).Displayed)
                {
                    for (var i = 1; i <= rowCountI; i++)
                    {
                        var xPathTemp = xPath + i.ToString() + "']";
                        if (_clsCommon.FindElement(webDriver, By.XPath(xPathTemp), _clsCommon.IntTimeOut).Displayed)
                        {
                            _clsCommon.FindElement(webDriver, By.XPath(xPathTemp), _clsCommon.IntTimeOut).Click();
                            //Check color of page navigation
                            VerifyColorObject(webDriver, strObjectName, xPathTemp, strData);
                            //Check the content of page navigation
                            VerifyText(webDriver, strObjectName, xPathTemp, i.ToString());
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }
        public bool CheckDataTableOfLetterTemplate(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            const string xPath = ".//*[@id='gridLetterTemplate']/nav/ul/li[..]/a[text()='";
            var bResult = false;
            var rowCount = CompareLetterTemplateTable(strData);
            var row = Convert.ToInt64(rowCount);
            var rowCountI = (int)row / 50;
            if (row % 50 > 0)
            {
                rowCountI += 1;
            }
            try
            {
                if (_clsCommon.FindElement(webDriver, By.XPath(oObjectXpath), _clsCommon.IntTimeOut).Displayed)
                {
                    for (var i = 1; i <= rowCountI; i++)
                    {
                        var xPathTemp = xPath + i.ToString() + "']";
                        if (_clsCommon.FindElement(webDriver, By.XPath(xPathTemp), _clsCommon.IntTimeOut).Displayed)
                        {
                            _clsCommon.FindElement(webDriver, By.XPath(xPathTemp), _clsCommon.IntTimeOut).Click();
                            //Thread.Sleep(4000);
                            //Check data on a page
                            bResult = CheckDataOnOnePage(webDriver, strData, i);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
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
        ///CheckDataOnApage() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Check result of searching on one "Letter Template" page
        ///Description: compare result table to searched text
        /// </summary>
        public bool CheckDataOnOnePage(IWebDriver webDriver, string strData, Int64 numberOfPage)
        {
            var bResult = false;
            var bCodeResultI = false;
            var bCodeResultD = false;
            var bNameResultI = false;
            var bNameResultD = false;
            const string xPath1 = "//*[@id='gridLetterTemplate']/table/tbody/tr[..]/td[1]";
            const string xPath2 = "//*[@id='gridLetterTemplate']/table/tbody/tr[..]/td[2]";
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(xPath1));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);

                if (oWebElements.Any())
                {
                    foreach (var item in oWebElements)
                    {
                        //check the interface
                        if (item.Text.Contains(strData))
                        {
                            bCodeResultI = true;
                        }
                        //check Database
                        var chkDb = CompareLetterTemplateTable(item.Text);
                        if (chkDb.Equals("1"))
                        {
                            bCodeResultD = true;
                        }
                    }
                    if (bCodeResultI == false)
                    {
                        int j = 0;
                        ReadOnlyCollection<IWebElement> oWebElements2;
                        do
                        {
                            oWebElements2 = webDriver.FindElements(By.XPath(xPath2));
                            j += SleepTime;
                            Thread.Sleep(SleepTime);
                        } while (!oWebElements2.Any() && j < _clsCommon.IntTimeSearch);
                        foreach (var item in oWebElements2)
                        {
                            //check the interface
                            if (item.Text.Contains(strData))
                            {
                                bNameResultI = true;
                            }
                            //check Database
                            var chkDb = CompareLetterTemplateTable(item.Text);
                            if (chkDb.Equals("1"))
                            {
                                bNameResultD = true;
                            }
                        }

                    }
                    string strStepDetails;
                    if ((bCodeResultI || bNameResultI) && (bCodeResultD || bNameResultD))
                    {
                        strStepDetails = "Passed. Data of page " + numberOfPage + " is matched.";
                        bResult = true;
                    }
                    else
                    {
                        strStepDetails = "Fail. Data of page " + numberOfPage + " is not matched.";
                    }
                    Console.WriteLine(strStepDetails);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
            return bResult;
        }

        public string CompareLetterTemplateTable(string paramVal)
        {
            using (var connection = new SqlConnection(_conn))
            {
                using (var cmd = connection.CreateCommand())
                {
                    string sqlCmd;
                    if (paramVal != "")
                    {
                        sqlCmd = "SELECT COUNT(*) FROM CM_LetterTemplate WHERE LetterTemplateCode like'%" + paramVal + "%'" +
                                          " OR LetterTemplateName like '%" + paramVal + "%'";
                    }
                    else
                    {
                        sqlCmd = "SELECT COUNT(*) FROM CM_LetterTemplate";
                    }
                    cmd.CommandText = sqlCmd;
                    cmd.CommandType = CommandType.Text;

                    connection.Open();
                    var countTemp = cmd.ExecuteScalar();
                    connection.Close();
                    return countTemp.ToString();
                }
            }
        }
        // Purpose: check data after insert into LetterTemplate Table
        public bool VerifyInsertIntoLetterTemplateTable(string parCode)
        {
            string strStepDetails = "";
            var bResult = false;
            try
            {
                using (var connection = new SqlConnection(_conn))
                {
                    using (var cmd = connection.CreateCommand())
                    {
                        if (!parCode.IsNullOrEmpty())
                        {
                            var sqlCmd = "SELECT COUNT(*) FROM CM_LetterTemplate WHERE LetterTemplateCode ='" + parCode + "'";
                            cmd.CommandText = sqlCmd;
                            cmd.CommandType = CommandType.Text;
                            connection.Open();
                            var countTemp = cmd.ExecuteScalar();
                            connection.Close();
                            if (countTemp.ToString() == "1")
                            {
                                strStepDetails = "Passed. Data inserted into CM_LetterTemplate table successfully!";
                                bResult = true;
                            }
                            else
                            {
                                strStepDetails = "Fail. Data did not insert into CM_LetterTemplate table";
                            }
                        }
                    }
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
        public string GetValue(string sproc, string paramVal)
        {
            using (var connection = new SqlConnection(_conn))
            {
                using (var cmd = connection.CreateCommand())
                {
                    connection.Open();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = @"" + sproc;
                    cmd.Parameters.AddWithValue("@KEY_ID", paramVal);
                    var k = cmd.ExecuteScalar();
                    connection.Close();
                    return k.ToString();
                }
            }
        }
        /// <summary>
        ///AssertUIandDatabaseTrue() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Compare data between User Interface and Database
        ///Description: 
        /// input:
        /// data on User Interface to be stored in OldData array
        /// data from Database to be got by SPWS_GET_SYSINFO_BY_KEYID stored procedure and store NewData array
        /// strData == Column Name 
        /// Output:
        /// Return true if those input is matched and vice versa
        /// </summary>
        public bool AssertUIandDatabaseTrue(string strObjectName, string strData)
        {
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                var valFromDb = GetValue("SPWS_GET_SYSINFO_BY_KEYID", strData);
                var valOld = OldData[strObjectName].ToLower();
                string strStepDetails;
                if (valFromDb.ToLower().Equals(valOld.ToLower()))
                {
                    strStepDetails = "Passed. [" + valOld.ToUpper() + "] value is saved to " + strData.ToUpper() + " field of Database.";
                    bResult = true;
                }
                else
                {
                    strStepDetails = "Fail. [" + valOld.ToUpper() + "] value differ to [" + valFromDb.ToUpper() + "] value in Database.";
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
        ///AssertUIandDatabaseFalse() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Compare data between User Interface and Database
        ///Description: 
        /// input:
        /// data on User Interface to be stored in OldData array
        /// data from Database to be got by SPWS_GET_SYSINFO_BY_KEYID stored procedure and store NewData array
        /// strData == Column Name  
        /// Output:
        /// Return true if those input is not matched and vice versa
        /// </summary>
        public bool AssertUIandDatabaseFalse(string strObjectName, string strData)
        {
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                var valFromDb = GetValue("SPWS_GET_SYSINFO_BY_KEYID", strData);
                var valOld = OldData[strObjectName].ToLower();
                string strStepDetails;
                if (valFromDb.ToLower().Equals(valOld.ToLower()))
                {
                    strStepDetails = "Fail. [" + valOld + "] value is saved to " + strData + " field of Database.";
                }
                else
                {
                    strStepDetails = "Passed. [" + valOld + "] value differ to [" + valFromDb + "] value in Database.";
                    bResult = true;
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

        public bool VerifyCheckBoxIsSelected(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            var isEnable = false;
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
                    isEnable = oWebElements[0].Selected;
                }
                if (isEnable)
                {
                    strStepDetails.AppendFormat("Passed. The {0} is checked.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT checked.", strObjectName);
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
        public bool VerifyCheckBoxIsNotSelected(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            var isEnable = false;
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
                    isEnable = oWebElements[0].Selected;
                }
                if (isEnable == false)
                {
                    strStepDetails.AppendFormat("Passed. The {0} is NOT checked.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is checked.", strObjectName);
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
        public bool ClearNewData()
        {
            var bResult = false;
            try
            {
                NewData.Clear();
                const string strStepDetails = "Passed. Clear data successfully.";
                bResult = true;
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
        public bool ClearOldData()
        {
            var bResult = false;
            const string strStepDetails = "Passed. Clear data successfully.";
            try
            {
                OldData.Clear();
                bResult = true;
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
        ///CaptureNewValue() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Setting data for public array NewData
        ///Description: use method GetAttribute("value") for setting data
        /// </summary>
        public bool CaptureNewValue(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
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
                    var value = oWebElements[0].GetAttribute("value").Trim();
                    if (NewData.ContainsKey(strObjectName))
                    {
                        NewData.Remove(strObjectName);
                    }
                    NewData.Add(strObjectName, value);
                    strStepDetails.AppendFormat("Passed. Capture the new value of {0} successfully.The new one is {1}",
                        strObjectName, value);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. The name of [{0}] is duplicated.", strObjectName);
                bResult = false;
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
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
        ///CaptureOldValue() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: Setting data for public array OldData
        ///Description: use method GetAttribute("value") for setting data
        /// </summary>
        public bool CaptureOldValue(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
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
                    var value = oWebElements[0].GetAttribute("value").Trim();
                    if (OldData.ContainsKey(strObjectName))
                    {
                        OldData.Remove(strObjectName);
                    }
                    OldData.Add(strObjectName, value);
                    strStepDetails.AppendFormat("Passed. Capture {0} of {1} successfully.", value, strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Object {0} is not displayed.", strObjectName);
                }

                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. Object {0} is duplicated.", strObjectName);
                bResult = false;
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
        ///CaptureOldValueOfCheckField() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Setting data for public array OldData
        ///Description: check that the object is displayed or not
        /// </summary>
        public bool CaptureOldValueOfCheckField(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                if (OldData.ContainsKey(strObjectName))
                {
                    OldData.Remove(strObjectName);
                }
                ReadOnlyCollection<IWebElement> element = webDriver.FindElements(By.XPath(oObjectXpath));
                if (element.Any())
                {
                    OldData.Add(strObjectName, true.ToString());
                    strStepDetails.AppendFormat("Passed. Capture the value {0} for {1} successfully", true, strObjectName);
                    bResult = true;
                }
                else
                {
                    OldData.Add(strObjectName, false.ToString());
                    strStepDetails.AppendFormat("Passed. Capture the value {0} for {1} successfully", false, strObjectName);
                    bResult = true;
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. The name of {0} is duplicated.", strObjectName);
                bResult = false;
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
        ///CaptureNewValueOfCheckField() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Setting data for public array OldData
        ///Description: check that the object is displayed or not
        /// </summary>
        public bool CaptureNewValueOfCheckField(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                if (NewData.ContainsKey(strObjectName))
                {
                    NewData.Remove(strObjectName);
                }
                ReadOnlyCollection<IWebElement> element = webDriver.FindElements(By.XPath(oObjectXpath));
                if (element.Any())
                {
                    NewData.Add(strObjectName, true.ToString());
                    strStepDetails.AppendFormat("Passed. Capture the value {0} for {1} successfully", true, strObjectName);
                    bResult = true;
                }
                else
                {
                    NewData.Add(strObjectName, false.ToString());
                    strStepDetails.AppendFormat("Passed. Capture the value {0} for {1} successfully", false, strObjectName);
                    bResult = true;
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. The name of {0} is duplicated.", strObjectName);
                bResult = false;
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
        ///CaptureOldData() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Setting data for public array OldData
        ///Description: use method Add for setting data
        /// </summary>
        public bool CaptureOldData(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
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
                    string value;
                    if (useragent.Contains("MSIE 8.0"))
                    {
                        value = oWebElements[0].Text.Trim(); // for IE 8    
                    }
                    else
                    {
                        value = oWebElements[0].GetAttribute("textContent").Trim(); // for IE 9 or above and Chrome.    
                        if (value.IsNullOrEmpty())
                        {
                            value = oWebElements[0].GetAttribute("value").Trim();
                        }
                    }

                    if (OldData.ContainsKey(strObjectName))
                    {
                        throw new DuplicateNameException();
                    }
                    OldData.Add(strObjectName, value);
                    strStepDetails.AppendFormat("Passed. Capture {0} of {1} successfully.", value, strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Object {0} is not displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. Object {0} is duplicated.", strObjectName);
                bResult = false;
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
        ///CaptureNewData() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Setting data for public array NewData
        ///Description: use method Add for setting data
        /// </summary>
        public bool CaptureNewData(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var strStepDetails = new StringBuilder();
            var bResult = false;
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
                    //var value = oWebElements[0].GetAttribute("textContent").Trim();
                    //if (value.IsNullOrEmpty())
                    //{
                    //    value = oWebElements[0].GetAttribute("value").Trim();
                    //}
                    string value;
                    if (useragent.Contains("MSIE 8.0"))
                    {
                        value = oWebElements[0].Text.Trim(); // for IE 8    
                    }
                    else
                    {
                        value = oWebElements[0].GetAttribute("textContent").Trim(); // for IE 9 or above and Chrome.    
                        if (value.IsNullOrEmpty())
                        {
                            value = oWebElements[0].GetAttribute("value").Trim();
                        }
                    }

                    if (NewData.ContainsKey(strObjectName))
                    {
                        throw new DuplicateNameException();
                    }
                    NewData.Add(strObjectName, value);
                    strStepDetails.AppendFormat("Passed. Capture {0} of {1} successfully.", value, strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. Object {0} is not displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (DuplicateNameException)
            {
                strStepDetails.AppendFormat("Fail. Object {0} is duplicated.", strObjectName);
                bResult = false;
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
        ///VerifyEquals() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Comparing old value to new value
        ///Description: use method Equals for comparing
        /// input: two arrays are newData and old data. Using  method CaptureNewData() to get data for array newData. 
        /// Using method CaptureOldData() to get data for array oldData.
        /// Output: it's true if the both value are same and vise versa.
        /// </summary>
        public bool VerifyEquals(IWebDriver webDriver, string strObjectName)
        {
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                string strStepDetails;
                //Compare two arrays
                var oldValue = OldData[strObjectName];
                var newValue = NewData[strObjectName];
                if (newValue.Equals(oldValue))
                {
                    strStepDetails = "Passed. The old value and the new value are " + newValue;
                    bResult = true;
                }
                else
                {
                    strStepDetails = "Fail. The old value and the new value are not equal. \n The old value is [" + oldValue + "] and the new value is [" + newValue + "]";
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
        ///VerifyContains() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Comparing old value to new value
        ///Description: use method Equals for comparing
        /// input: two arrays are newData and old data. Using  method CaptureNewData() to get data for array newData. 
        /// Using method CaptureOldData() to get data for array oldData.
        /// Output: it's true if the both value are same and vise versa.
        /// </summary>
        public bool VerifyContains(IWebDriver webDriver)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                //Compare two arrays
                // ReSharper disable once PossibleNullReferenceException
                var newData = NewData.Select(x => x.Value).FirstOrDefault();
                // ReSharper disable once PossibleNullReferenceException
                var oldData = OldData.Select(x => x.Value).FirstOrDefault();
                if (newData.Contains(oldData))
                {
                    strStepDetails.AppendFormat("Passed. The old value is a part of the new value");
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The old value is NOT a part of new value. \n The old value is [" + OldData.ToString() + "] and the new value is [" + NewData.ToString() + "]");
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
        ///VerifyGreater() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Comparing old value to new value
        ///Description:
        /// input: two arrays are newData and old data. Using  method CaptureNewData() to get data for array newData. 
        /// Using method CaptureOldData() to get data for array oldData.
        /// Output: it's true if the new value is greater than the old one.
        /// </summary>
        public bool VerifyGreater(IWebDriver webDriver, string strObjectName)
        {
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                string strStepDetails;
                //Compare two arrays
                int oldValue = int.Parse(OldData[strObjectName]);
                int newValue = int.Parse(NewData[strObjectName]);
                if (newValue > oldValue)
                {
                    strStepDetails = "Passed. [" + newValue + "] is greater than [" + oldValue + "]";
                    bResult = true;
                }
                else
                {
                    strStepDetails = "Fail. The old value and the new data are not the equal. \n The old value is [" + oldValue + "] and the new value is [" + newValue + "]";
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
        ///VerifyNotEqual() keyword description:
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: July-2015
        ///Purpose: Comparing old value to new value
        ///Description:
        /// input: two arrays are newData and old data. Using  method CaptureNewData() to get data for array newData. 
        /// Using method CaptureOldData() to get data for array oldData.
        /// Output: it's true if the new value is not equal to the old one.
        public bool VerifyNotEqual(IWebDriver webDriver, string strObjectName)
        {
            var bResult = false;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                string strStepDetails;
                //Compare two arrays
                var oldValue = OldData[strObjectName];
                var newValue = NewData[strObjectName];
                if (newValue != oldValue)
                {
                    strStepDetails = "Passed. [" + newValue + "] is not equal to [" + oldValue + "]";
                    bResult = true;
                }
                else
                {
                    strStepDetails = "Fail. The old value and the new data are the equal. \n It is " + oldValue;
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
        ///VerifyBackGroundColorObject() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript) 
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is displayed and vise versa
        /// </summary>
        public bool VerifyBackGroundColorObject(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            var bResult = false;
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
                    var color = oWebElements[0].GetCssValue("background-color");
                    var length = color.LastIndexOf(')') - 4;
                    var temp = color.Substring(5, length);
                    var array = temp.Split(',');
                    var myColor = Color.FromArgb(int.Parse(array[0]), int.Parse(array[1]), int.Parse(array[2]));
                    //return color code such as: ff0000
                    var hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
                    if (hex == strData)
                    {
                        strStepDetails.AppendFormat("Passed. The background color of {0} is {1}", strObjectName, hex);
                        bResult = true;
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. The actual background color is {1}, but the expected background color is {0}", strData, hex);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is not displayed.", strObjectName);
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
        ///VerifyColorObject() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript) 
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is displayed and vise versa
        /// </summary>
        public bool VerifyColorObject(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            var bResult = false;
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
                    if (useragent.Contains("MSIE 8.0")) // for IE 8    
                    {
                        bResult = true;
                        strStepDetails.AppendFormat("Passed. The color of {0} is true.", strObjectName);
                    }
                    else //for IE 9, 10,11, Chrome
                    {
                        var color = oWebElements[0].GetCssValue("Color");
                        var length = color.LastIndexOf(')') - 4;
                        var temp = color.Substring(4, length);
                        var array = temp.Split(',');
                        var myColor = Color.FromArgb(int.Parse(array[0]), int.Parse(array[1]), int.Parse(array[2]));
                        //return color code such as: ff0000
                        var hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
                        if (hex == strData)
                        {
                            strStepDetails.AppendFormat("Passed. The color of {0} is {1}", strObjectName, hex);
                            bResult = true;
                        }
                        else
                        {
                            strStepDetails.AppendFormat("Fail. The result color of {0} is {1}, but the expected color is {2}", strObjectName, hex, strData);
                        }
                        
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is not displayed.", strObjectName);
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
        ///VerifyAnObjectPresent() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Updated Date: Feb-2016
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript) 
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is displayed and against
        /// </summary>
        public bool VerifyAnObjectPresent(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any() && oWebElements[0].Displayed)
                {
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strObjectName);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
            return bResult;
        }
        /// <summary>
        ///VerifyAnObjectNotPresent() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        /// Updated Feb-2016
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript)
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is not displayed and against
        /// </summary>
        public bool VerifyAnObjectNotPresent(IWebDriver webDriver, string strObjectName, string oObjectXpath)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any() && oWebElements[0].Displayed)
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strObjectName);
                }
                else
                {
                    bResult = true;
                    strStepDetails.AppendFormat("Passed. [{0}] is not displayed.", strObjectName);
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
        ///VerifySearchFieldPresent() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript)  on "Customer Search" client page
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is displayed and against
        /// </summary>
        public bool VerifySearchFieldPresent(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                var oObjectXpath = string.Format("//*[@id='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Passed. [{0}] is displayed.", strData);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
            return bResult;
        }
        /// <summary>
        ///VerifySearchFieldNotPresent() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Apr-2015
        ///Purpose: Check an object is hidden or not (Developer uses JavaScript)  on "Customer Search" client page
        ///Description: use method Displayed for checking
        /// Output: it's true if the object is not displayed and against
        /// </summary>
        public bool VerifySearchFieldNotPresent(IWebDriver webDriver, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                var oObjectXpath = string.Format("//*[@id='{0}']", strData);
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is displayed.", strData);
                    
                }
                else
                {
                    strStepDetails.AppendFormat("Passed. [{0}] is not displayed.", strData);
                    bResult = true;
                }
                Console.WriteLine(strStepDetails);
                Log.Info(strStepDetails);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
                return false;
            }
            return bResult;
        }
        /// <summary>
        ///ClickOnItemOfDropDown() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Mar-2015
        ///Purpose: select an item of list box
        ///Description: 
        ///input data includes: item name and xPath of combo box, which has subTag //option
        ///out out: selected item
        /// </summary>
        public bool ClickOnItemOfDropDown(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            int i = 0;
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                //Get the hold of the dropdown box by xPath or Nam
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    //Place the drop down into selectElemnet
                    SelectElement clickThisitem = new SelectElement(oWebElements[0]);
                    //Select the Item from dropdown by Text or index
                    clickThisitem.SelectByText(strData);
                    strStepDetails.AppendFormat("Passed. Choose {0} of {1} successfully.", strData, strObjectName);
                    bResult = true;
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} of {1} is NOT displayed.", strData, strObjectName);
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
        ///VerifyItemOfCombo() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Nov-2015
        ///Purpose: select an item of combo box
        ///Description: 
        ///input data includes: item name and xPath of combo box, which has subTag //option
        ///out out: selected item
        public bool VerifyItemOfCombo(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    int j = 0;
                    ReadOnlyCollection<IWebElement> oWebSubElements;
                    do
                    {
                        oWebSubElements = webDriver.FindElements(By.XPath(oObjectXpath + "//option"));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebSubElements.Any() && j < _clsCommon.IntTimeSearch);

                    if (oWebSubElements.Any())
                    {
                        var dpListCount = oWebSubElements.Count;
                        if (strData == "Null")
                        {
                            oWebSubElements[0].Click();
                            bResult = true;
                        }
                        for (var k = 0; k < dpListCount; k++)
                        {
                            if (oWebSubElements[k].Text == strData)
                            {
                                bResult = true;
                                break;
                            }
                        }
                    }
                    if (bResult)
                    {
                        strStepDetails.AppendFormat("Passed. The {0} belongs to {1}.", strData, strObjectName);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. The {0} does not belong to {1}.", strData, strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strObjectName);
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
        ///ClickOnItemOfCombo() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Mar-2015
        ///Purpose: select an item of combo box
        ///Description: 
        ///input data includes: item name and xPath of combo box, which has subTag //option
        ///out out: selected item
        /// </summary>
        public bool ClickOnItemOfCombo(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    int j = 0;
                    ReadOnlyCollection<IWebElement> oWebSubElements;
                    do
                    {
                        oWebSubElements = webDriver.FindElements(By.XPath(oObjectXpath + "//option"));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebSubElements.Any() && j < _clsCommon.IntTimeSearch);

                    if (oWebSubElements.Any())
                    {
                        var dpListCount = oWebSubElements.Count;
                        if (strData == "Null")
                        {
                            oWebSubElements[0].Click();
                            bResult = true;
                        }
                        for (var k = 0; k < dpListCount; k++)
                        {
                            if (oWebSubElements[k].Text.Trim()==strData)
                            {
                                oWebSubElements[k].Click();
                                bResult = true;
                                break;
                            }
                        }
                    }
                    if (bResult)
                    {
                        strStepDetails.AppendFormat("Passed. Choose {0} of {1} successfully.", strData, strObjectName);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. There is not {0} of {1}.", strData, strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
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
        public bool ClickOnItemOfCombo2(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            bool bResult = false;
            var strStepDetails = new StringBuilder();
            IJavaScriptExecutor js = webDriver as IJavaScriptExecutor;
            //Actions action = new Actions(webDriver);
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    int j = 0;
                    ReadOnlyCollection<IWebElement> oWebSubElements;
                    do
                    {
                        oWebSubElements = webDriver.FindElements(By.XPath(oObjectXpath + "//option"));
                        j += SleepTime;
                        Thread.Sleep(SleepTime);
                    } while (!oWebSubElements.Any() && j < _clsCommon.IntTimeSearch);

                    if (oWebSubElements.Any())
                    {
                        var dpListCount = oWebSubElements.Count;
                        if (strData == "Null")
                        {
                            oWebSubElements[0].Click();
                            bResult = true;
                        }
                        for (var k = 0; k < dpListCount; k++)
                        {
                            if (oWebSubElements[k].Text.Trim() == strData)
                            {
                                //oWebSubElements[k].Click();
                                //action.MoveToElement(oWebSubElements[k]).Build().Perform();
                                //if (js != null) js.ExecuteScript("arguments[0].click()", oWebSubElements[k]);
                                //action.MoveToElement(oWebSubElements[k]).Build().Perform();
                                oWebSubElements[k].Click();
                                oWebSubElements[k].SendKeys("Keys.Enter");
                                bResult = true;
                                break;
                                //action.Click();
                                //action.SendKeys("Keys.Enter");
                                //bResult = true;
                                //break;
                            }
                        }
                    }
                    if (bResult)
                    {
                        strStepDetails.AppendFormat("Passed. Choose {0} of {1} successfully.", strData, strObjectName);
                    }
                    else
                    {
                        strStepDetails.AppendFormat("Fail. There is not {0} of {1}.", strData, strObjectName);
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strData);
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
        ///ClickOnItemOfGroupCombo() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: May-2015
        ///Purpose: select an item of the group of combo box
        ///Description: 
        ///input data includes: item name and xPath of combo box, which has subTag //optgroup//option
        ///out out: selected item
        /// </summary>
        public bool ClickOnItemOfGroupCombo(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait); 
            try
            {
                var nameArr = strData.Split('|');
                var element = _clsCommon.FindElement(webDriver, By.XPath(oObjectXpath + "/optgroup[@label='" + nameArr[0] + "']"), _clsCommon.IntTimeOut);
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = element.FindElements(By.XPath("//option"));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    for (var j = 0; j < oWebElements.Count; j++)
                    {
                        if (oWebElements[j].Text.Trim() == nameArr[1])
                        {
                            oWebElements[j].Click();
                            //oWebElements[j].SendKeys(nameArr[1]);
                            bResult = true;
                            break;
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} is NOT displayed.", strObjectName);
                }
                if (bResult)
                {
                    strStepDetails.AppendFormat("Passed. Choose {0} of {1} successfully.", nameArr[1], strObjectName);
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. The {0} of {1} is NOT displayed.", nameArr[1], strObjectName);
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
        /// <summary>
        ///VerifyDefaultTextInCombo() 
        ///Written by: Viet.LeMinh2@Harveynash.vn
        ///Date: Mar-2015
        ///Purpose: Verify default value of combo box
        ///Description: 
        ///input : expected result  and xPath of combo box
        ///out put: Return true if actual result equas to expected result and otherwhile return false value
        /// </summary>
        public bool VerifyDefaultTextInCombo(IWebDriver webDriver, string strObjectName, string oObjectXpath, string strData)
        {
            var bResult = false;
            var strStepDetails = new StringBuilder();
            Thread.Sleep(_clsCommon.IntTimeWait);
            try
            {
                int i = 0;
                ReadOnlyCollection<IWebElement> oWebElements;
                do
                {
                    oWebElements = webDriver.FindElements(By.XPath(oObjectXpath));
                    i += SleepTime;
                    Thread.Sleep(SleepTime);
                } while (!oWebElements.Any() && i < _clsCommon.IntTimeSearch);
                if (oWebElements.Any())
                {
                    var selectedValue = new SelectElement(oWebElements[0]);
                    var txtMessage = selectedValue.SelectedOption.Text;
                    if (strData.IsNullOrEmpty() && strData.IsNullOrEmpty())
                    {
                        //strStepDetails.AppendFormat("Fail. There isn't any Test Data");
                        bResult = true;
                        strStepDetails = strStepDetails.AppendFormat("Passed. The data are matched.");
                    }
                    else
                    {
                        if (txtMessage != "")
                        {
                            if (txtMessage.Contains(strData.Trim()))
                            {
                                bResult = true;
                                strStepDetails = strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. Data are not matched. Expected Result is {0} but Actual Result is {1}", strData.Trim(), txtMessage.Trim());
                            }
                        }
                        else
                        {
                            var value = selectedValue.SelectedOption.Text;
                            if (value.Equals(strData))
                            {
                                bResult = true;
                                strStepDetails.AppendFormat("Passed. The data are matched.");
                            }
                            else
                            {
                                strStepDetails.AppendFormat("Fail. Data are not matched. Expected Result is {0} but Actual Result is {1}", strData.Trim(), value);
                            }
                        }
                    }
                }
                else
                {
                    strStepDetails.AppendFormat("Fail. [{0}] is not displayed.", strObjectName);
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
    }
}
