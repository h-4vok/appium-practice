using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;

namespace AppiumPractice
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            // Launch the calculator app
            //DesiredCapabilities appCapabilities = new DesiredCapabilities();
            //appCapabilities.SetCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");
            //var CalculatorSession = new RemoteWebDriver(new Uri(" http://127.0.0.1:4723/"), appCapabilities);
            
            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");

            var CalculatorSession = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);
            Assert.IsNotNull(CalculatorSession);

            CalculatorSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            CalculatorSession.FindElementByName("One").Click();
            CalculatorSession.FindElementByName("Plus").Click();
            CalculatorSession.FindElementByName("Seven").Click();
            CalculatorSession.FindElementByName("Equals").Click();

            var CalculatorResults = CalculatorSession.FindElementByAccessibilityId("CalculatorResults");
            Assert.AreEqual("Display is 8", CalculatorResults.Text);

            CalculatorSession.Close();
        }

        [TestMethod]
        public void TestMethod2()
        {
            // lets do it with M4
            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", @"C:\Program Files (x86)\M4Solutions\M4Solutions.exe");

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);
            Assert.IsNotNull(driver);

            var appWindow = driver.FindElementByAccessibilityId("MDIForm1");
            var projectCenterWindow = appWindow.FindElementByAccessibilityId("ProjectCenter");

            // lets open a project (a bit!)
            var pickProjectButton = appWindow.FindElementByAccessibilityId("btnPickProject");
            pickProjectButton.Click();

            var openDialog = appWindow.FindElementByName("Open");
            var fileNameTextBox = openDialog.FindElementByAccessibilityId("1148");

            fileNameTextBox.SendKeys(@"E:\clients\mapcom\M4 Regression Test Automation\Test Project\Diversicom\ARVIG.prj");
            
            var openButton = openDialog.FindElementByAccessibilityId("1");
            openButton.Click();

            driver.Close();
        }

        [TestMethod]
        public void TestMethod3()
        {
            var firefox = @"C:\Program Files\Mozilla Firefox\firefox.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", firefox);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var searchBox = driver.FindElementByName("Search with Google or enter address");
            searchBox.SendKeys("www.google.com.ar");
            searchBox.SendKeys(Keys.Enter);
        }
    }
}
