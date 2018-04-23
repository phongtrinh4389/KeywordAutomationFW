using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;
using log4net;
using NPOI.Util;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;

namespace Selenium_Automated_Testing.Utilities
{
    public class Common
    {
        //	Init the parameter for write log
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        public static string StrTimeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        public string StrAppUnderTest = "";
        public string StrRootPath = "";
        public string StrBrowerPath = "";
        public string StrHtmlPath = "";
        public string StrDataPath = "";
        public string StrTestResultPath = "";
        public string StrHtmlReportPath = "";
        public string StrHtmlReportFile = "";
        public string StrDataFile = "";
        public string StrTestCaseFile = "";
        public string StrResultFile = "";
        public string StrSuiteSheet = "";
        public string RefProperties = "ref.properties";
        public string TableProperties = "table.properties";
        public string OrProperties = "or.properties";
        public string StrLogFile = "";
        public string StrSheetNameUrl = "";
        public string StrChromepath = "";
        // ReSharper disable once InconsistentNaming
        public string strIEPath = "";
        public string StrRootProjectPath;
        public string StrPasswordEmail = "";
        public static string ConnectionString;
        public string StrBrowserType = "";
        private FirefoxProfile _ffp;
        public string StrFromAddress = "";
        public string StrToAddress = "";
        public string StrSmtp = "";
        public Int32 IntPort;
        public int IntColumnOfResultTestCaseDetails;
        public int IntColumnOfResultTestSuiteDetails;
        public int IntColumnOfPass;
        public int IntTimeOut;
        public int IntTimeSearch;
        public int IntTimeWait;
        public int IntClickWait;
        public Common()
        {
            // Get the path of ROOT            
            string strRootPath = Directory.GetCurrentDirectory();
            StrRootProjectPath = strRootPath.Substring(0, strRootPath.Length - 9);
            LoadRefProperties();
        }
     public void SendMail()
     {
         try
         {
             //SmtpClient smtpServer = new SmtpClient()
             //{
             //    Host = StrSmtp,
             //    Port = IntPort,
             //    UseDefaultCredentials = false,
             //    DeliveryMethod = SmtpDeliveryMethod.Network,
             //    EnableSsl = true,
             //    Credentials = new NetworkCredential(StrFromAddress, StrPasswordEmail)
             //};

             //// ReSharper disable once UseObjectOrCollectionInitializer
             //MailMessage mail = new MailMessage();
             //mail.From = new MailAddress(StrFromAddress);
             //mail.To.Add(StrToAddress);
             //mail.BodyEncoding = Encoding.UTF8;
             //mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
             //mail.IsBodyHtml = true;

             //mail.Subject = "The result of Automation Testing";
             //mail.Body = "Dear you, <br> Please see the result of testing in the attachment File!<br> Sincerely, <br> Automation Team";

             //var attachment = new Attachment(StrResultFile);
             //mail.Attachments.Add(attachment);
             //smtpServer.Send(mail);
         }
         catch (Exception e)
         {
             string error = String.Format("Send email got error. {0}", e);
             Console.WriteLine(error);
             Log.Error(error);
         }   
         
     }
        //Purpose: This function will read ref.properties file and set values for variables in class
        public void LoadRefProperties()
        {
            try
            {               
                Dictionary<string, string> dictionary = ReadPropertiesFile(RefProperties);

                StrDataPath = StrRootPath + dictionary["strDataPath"];
                StrTestResultPath = StrRootProjectPath + dictionary["strTestResultPath"];
                StrHtmlReportPath = StrRootProjectPath + dictionary["strHTMLReportPath"];
                StrHtmlReportFile = StrHtmlReportPath + dictionary["strHTMLReportFile"];
                StrDataFile = StrDataPath + dictionary["strDataFile"];
                StrTestCaseFile = StrDataPath + dictionary["strTestCasesFile"];
                StrLogFile = StrRootPath + dictionary["strLogFile"];
                StrSheetNameUrl = dictionary["strSheetURL"];
                StrSuiteSheet = dictionary["strSuiteSheet"];
                IntColumnOfResultTestCaseDetails = int.Parse(dictionary["intColumnOfResultTestCaseDetails"]);
                IntColumnOfResultTestSuiteDetails = int.Parse(dictionary["intColumnOfResultTestSuiteDetails"]);
                IntColumnOfPass = int.Parse(dictionary["intColumnOfPass"]);
                IntTimeOut = int.Parse(dictionary["intTimeout"]);
                IntTimeSearch = int.Parse(dictionary["intTimeSearch"]);
                IntTimeWait = int.Parse(dictionary["intTimeWait"]);
                IntClickWait = int.Parse(dictionary["intClickWait"]);
                StrResultFile = StrTestResultPath + "Result_" + StrTimeStamp + dictionary["strDataFile"] ;
                StrBrowserType = dictionary["strBrowserType"];
                StrChromepath = StrRootProjectPath + dictionary["StrChromepath"];
                strIEPath = StrRootProjectPath + dictionary["strIEPath"];
                StrFromAddress = dictionary["strFromAddress"];
                StrToAddress = dictionary["strToAddress"];
                StrSmtp = dictionary["strSMTP"];
                IntPort = int.Parse(dictionary["intPort"]);
                ConnectionString = dictionary["strConnectionString"];
                StrPasswordEmail = dictionary["strPassword"];
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
        }

        public string LoadTableParameter(string strTableName)
        {
            string strTablbeXpath = "";
            try
            {
                Dictionary<string, string> dictionary = ReadPropertiesFile(TableProperties);
                strTablbeXpath = dictionary[strTableName];
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return strTablbeXpath;
        }

        //Purpose: This function will read Properties File then return the result in Dictionary objecy
        public Dictionary<string, string> ReadPropertiesFile(string fileName)
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            try
            {
                string strPropertiesFile = StrRootPath + fileName;
                foreach (string line in File.ReadAllLines(strPropertiesFile, Encoding.Default))
                {
                    if ((!string.IsNullOrEmpty(line)) &&
                        (!line.StartsWith(";")) &&
                        (!line.StartsWith("#")) &&
                        (!line.StartsWith("'")) &&
                        (line.Contains("=")))
                    {
                        // ReSharper disable once StringIndexOfIsCultureSpecific.1
                        int index = line.IndexOf("=");
                        string key = line.Substring(0, index).Trim();
                        string value = line.Substring(index + 1).Trim();

                        if ((value.StartsWith("\"") && value.EndsWith("\"")) ||
                            (value.StartsWith("'") && value.EndsWith("'")))
                        {
                            value = value.Substring(1, value.Length - 2);
                        }
                        if (!dictionary.ContainsKey(key))
                            dictionary.Add(key, value);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Log.Error(e.ToString());
            }
            return dictionary;
        }

        public IWebDriver StartBrowser(string strBrowserType)
        {

            IWebDriver driver = null;
            switch (strBrowserType.ToLower())
            {
                case "firefox":
                    {
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        _ffp = new FirefoxProfile();
                        _ffp.AcceptUntrustedCertificates = true;
                        driver = new FirefoxDriver(_ffp);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMilliseconds(IntTimeOut));
                        break;
                    }
                case "iexplore":
                    {                        
                        // ReSharper disable once UseObjectOrCollectionInitializer
                        InternetExplorerOptions opts = new InternetExplorerOptions();
                        opts.IgnoreZoomLevel = true;
                        opts.IntroduceInstabilityByIgnoringProtectedModeSettings=true;
                        opts.EnableNativeEvents = true;
                        driver = new InternetExplorerDriver(strIEPath, opts);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMilliseconds(IntTimeOut));
                        break;
                    }
                case "chrome":
                    {
                        driver = new ChromeDriver(StrChromepath);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMilliseconds(IntTimeOut));
                        break;
                    }
            }
            return driver;
        }

        public void NavigateUrl(IWebDriver webDriver, string strUrl)
        {
            try
            {
                webDriver.Navigate().GoToUrl(strUrl);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        
        /// Purpose: This function find and return the element we want to get
        public IWebElement FindElement(IWebDriver webDriver, By by, int iTimeOut)
        {
            const int iSleepTime = 2000;
            for (int i = 0; i < iTimeOut; i += iSleepTime)
            {
                ReadOnlyCollection<IWebElement> oWebElements = webDriver.FindElements(by);
                if (oWebElements.Any())
                {
                    return oWebElements[0];
                }
                try
                {
                    Thread.Sleep(iSleepTime);
                    //Console.WriteLine("Waited for {0} milliseconds.{1}", i + iSleepTime, @by);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            // Can't find 'by' element. Therefore throw an exception.
            String sException = String.Format("Can't find {0} after {1} milliseconds.", by, iTimeOut);
            throw new RuntimeException(sException);
        }

        /// Purpose: This function will wait an object present until timeup.
        public void WaitAnElementPresent(IWebDriver webDriver, By by, int iTimeOut)
        {
            int iTotal = 0;
            const int iSleepTime = 5000;
            while (iTotal < iTimeOut)
            {
                ReadOnlyCollection<IWebElement> oWebElements = webDriver.FindElements(by);
                if (oWebElements.Count > 0)
                    return;
                try
                {
                    Thread.Sleep(iSleepTime);
                    iTotal = iTotal + iSleepTime;
                    Console.WriteLine("Waited for {0} in milliseconds {1}", iTotal, @by);
                }
                catch (Exception ex)
                {
                    throw new RuntimeException(ex);
                }
            }
        }

    }
}
