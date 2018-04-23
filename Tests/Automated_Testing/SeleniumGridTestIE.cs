using MbUnit.Framework;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using Selenium_Automated_Testing.Utilities;
using Worksmart_Automated_Testing.Utilities;

namespace Demo
{
    [Parallelizable]
    [TestFixture()]
    // ReSharper disable once InconsistentNaming
    public class SeleniumGridTestIE
    {
        private IWebDriver _driver;
        private readonly Controller _clsController = new Controller();
        private readonly Common _clsCommon = new Common();
        [SetUp]
        public void Init()
        {
            var capabilities = DesiredCapabilities.InternetExplorer();
            capabilities.SetCapability(CapabilityType.BrowserName, "Internet Explorer");
            //capabilities.SetCapability(CapabilityType.BrowserName, "iexplorer");
            capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
            //Hub & node is a machine: use localhost
            _driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), capabilities);
            //node is a remote machine: ip of node
            //_driver = new RemoteWebDriver(new Uri("http://172.16.21.19:5555/wd/hub"), capabilities);

        }
        [Test, Timeout(99999)]
        public void TestCase()
        {
            _clsController.RunTestSuite(_driver);
            _clsCommon.SendMail();
        }

        [TearDown]
        public void Cleanup()
        {
            _driver.Quit();
        }
    }
}