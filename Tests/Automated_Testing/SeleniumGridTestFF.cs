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
    public class SeleniumGridTestFireFox
    {
        private IWebDriver _driver;
        private readonly Controller _clsController = new Controller();
        private readonly Common _clsCommon = new Common();
        [SetUp]
        public void Init()
        {
            var capabilities = DesiredCapabilities.Firefox();
            capabilities.SetCapability(CapabilityType.BrowserName, "firefox");
            capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
            //capabilities.SetCapability(CapabilityType.Version, "31.0");
            _driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), capabilities);
            //IP=19
            //driver = new RemoteWebDriver(new Uri("http://172.16.21.19:5557/wd/hub"), capabilities);
            //IP=36
            //driver = new RemoteWebDriver(new Uri("http://172.16.21.36:5557/wd/hub"), capabilities);
            //IP=20.21
            //driver = new RemoteWebDriver(new Uri("http://172.16.20.21:5557/wd/hub"), capabilities);
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