using System;
using MbUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using Selenium_Automated_Testing.Utilities;
using Worksmart_Automated_Testing.Utilities;

namespace Demo
{
    [TestFixture()]
    [Parallelizable]
    public class BrowserStackChrome
    {
        private IWebDriver _driver;
        private readonly Controller _clsController = new Controller();
        private readonly Common _clsCommon = new Common();
        [SetUp]
        public void Init()
        {
            DesiredCapabilities caps = DesiredCapabilities.Chrome();
            caps.SetCapability("browserstack.local", "true");
            caps.SetCapability("browserstack.user", "vietle3");
            caps.SetCapability("browserstack.key", "Apph7vqqDbAadTFLuQnq");
            caps.SetCapability("browser", "Chrome");
            caps.SetCapability("browser_version", "44.0");
            caps.SetCapability("os", "Windows");
            caps.SetCapability("os_version", "7");
            caps.SetCapability("resolution", "1024x768");
            caps.SetCapability("acceptSslCerts", "true");
            caps.SetCapability("browserstack.idleTimeout", "300");
            //_driver = new RemoteWebDriver(new Uri("http://vietle3.browserstack.com"), caps);
            _driver = new RemoteWebDriver(new Uri("http://hub.browserstack.com/wd/hub/"), caps);
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