using System;
using MbUnit.Framework;
using OpenQA.Selenium;
using Selenium_Automated_Testing.Utilities;
using Worksmart_Automated_Testing.Utilities;
namespace Demo
{
    [TestFixture]
    public class TestLocal
    {
        private IWebDriver _driver;
        private readonly Controller _clsController = new Controller();
        private readonly Common _clsCommon = new Common();
        
        //	Init the parameter for write log
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        [SetUp]
        public void Init()
        {
            try
            {
                string strBrowsertype = "";
                strBrowsertype = _clsController.GetBrowsertype();
                _driver = _clsCommon.StartBrowser(strBrowsertype);
                _driver.Manage().Window.Maximize();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
        }
        [Test, Timeout(99999)]
        public void RunTestSuite()
        {
            try
            {
                _clsController.RunTestSuite(_driver);
                _clsCommon.SendMail();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }

        }
        [TearDown]
        public void TearDown()
        {
            try
            {
                _driver.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Log.Error(ex.ToString());
            }
        }
    }
}
