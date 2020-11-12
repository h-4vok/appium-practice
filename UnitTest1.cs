using System;
using System.Drawing;
using Applitools.Images;
using System.Buffers.Text;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.ImageComparison;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using OpenQA.Selenium.Interactions;

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
            var pickProjectButton = projectCenterWindow.FindElementByAccessibilityId("btnPickProject");
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


            driver.Close();

        }

        [TestMethod]
        public void TestMethod4()
        {
            var notepad = @"C:\Windows\System32\notepad.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", notepad);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var eyes = new Eyes();
            eyes.ApiKey = "lFMnO9e1nbFQEdhebrDmhEBOqB7jXUXJGogj103lRWAzM110";
            eyes.SetAppEnvironment("Windows 10", null);

            try
            {
                eyes.Open("AppiumPractice", "TestMethod4");
                Thread.Sleep(3000);
                var currentAppScreenshot = driver.GetScreenshot();

                eyes.CheckImage(currentAppScreenshot.AsByteArray, "Notepad Just Opened");

                eyes.Close();
            }
            finally
            {
                eyes?.AbortIfNotClosed();
                driver.Quit();
            }
        }

        [TestMethod]
        public void TestMethod5()
        {
            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", @"C:\Program Files (x86)\M4Solutions\M4Solutions.exe");

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            var projectToOpen = @"E:\clients\mapcom\M4 Regression Test Automation\Test Project\Diversicom\ARVIG.prj";
            System.Windows.Forms.Clipboard.SetText(projectToOpen);

            var projectCenterWindow = driver.FindElementByAccessibilityId("ProjectCenter");
            var btnPickProject = projectCenterWindow.FindElementByAccessibilityId("btnPickProject");

            btnPickProject.Click();

            var openDialog = projectCenterWindow.FindElementByClassName("#32770");

            var fileTextBox = openDialog.FindElementByAccessibilityId("1148");

            fileTextBox.SendKeys(projectToOpen);

            var openButton = openDialog.FindElementByAccessibilityId("1");
            openButton.Click();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

            var mainWindow = driver.FindElementByAccessibilityId("MDIForm1");
            wait.Until(ExpectedConditions.TextToBePresentInElement(mainWindow, "M4 - ARVIG"));
            wait.Until(ExpectedConditions.ElementExists(By.Name("COUNTIES.msi")));

            var eyes = new Eyes();
            eyes.ApiKey = "lFMnO9e1nbFQEdhebrDmhEBOqB7jXUXJGogj103lRWAzM110";
            eyes.SetAppEnvironment("Windows 10", null);

            try
            {
                eyes.Open("AppiumPractice", "TestMethod5");
                Thread.Sleep(3000);
                var currentAppScreenshot = driver.GetScreenshot();

                eyes.CheckImage(currentAppScreenshot.AsByteArray, "Project Finished Loading");

                eyes.Close();
            }
            finally
            {
                eyes?.AbortIfNotClosed();
                driver.Quit();
            }
        }

        [TestMethod]
        public void TestMethod6()
        {
            var apikey = ConfigurationManager.AppSettings["APPLITOOLS_API_KEY"];
            Console.WriteLine(apikey);
        }

        [TestMethod]
        public void TestMethod7()
        {
            var src = Image.FromFile(Path.Combine(Environment.CurrentDirectory, "PatternTestNotepad.png"));
            var cropRect = new Rectangle(5, 5, 30, 10);
            var target = new Bitmap(cropRect.Width, cropRect.Height);

            using (var g = Graphics.FromImage(target))
            {
                g.DrawImage(src, new Rectangle(0, 0, target.Width, target.Height), cropRect, GraphicsUnit.Pixel);
            }

            target.Save(Path.Combine(Environment.CurrentDirectory, "pija.png"));
        }

        private byte[] CropScreenshot(Screenshot sshot)
        {
            using (var stream = new MemoryStream(sshot.AsByteArray))
            {
                var src = Image.FromStream(stream);
                var cropRect = new Rectangle(5, 5, 30, 10);
                var target = new Bitmap(cropRect.Width, cropRect.Height);

                using (var g = Graphics.FromImage(target))
                {
                    g.DrawImage(src, new Rectangle(0, 0, target.Width, target.Height), cropRect, GraphicsUnit.Pixel);
                    g.Flush();
                }

                var converter = new ImageConverter();
                var bytes = converter.ConvertTo(target, typeof(byte[])) as byte[];
                return bytes;
            }

        }

        private byte[] CropScreenshotWithRect(Screenshot sshot, WindowsElement we)
        {
            using (var stream = new MemoryStream(sshot.AsByteArray))
            {
                var src = Image.FromStream(stream);
                var target = new Bitmap(we.Size.Width, we.Size.Height);
                var srcRect = we.Rect;
                var destRect = new Rectangle(0, 0, we.Size.Width, we.Size.Height);

                using (var g = Graphics.FromImage(target))
                {
                    g.DrawImage(src, destRect, srcRect, GraphicsUnit.Pixel);
                    g.Flush();
                }

                var converter = new ImageConverter();
                var bytes = converter.ConvertTo(target, typeof(byte[])) as byte[];
                return bytes;
            }


        }

        public byte[] CropScreenshotOf2(byte[] screenshot, Rectangle rectangle)
        {
            using (var stream = new MemoryStream(screenshot))
            {
                var src = Image.FromStream(stream);
                var srcRect = rectangle;
                var destRect = new Rectangle(0, 0, srcRect.Width, src.Height);
                var target = new Bitmap(srcRect.Width, srcRect.Height);

                using (var g = Graphics.FromImage(target))
                {
                    g.DrawImage(src, destRect, srcRect, GraphicsUnit.Pixel);
                    g.Flush();
                }
                var converter = new ImageConverter();
                var bytes = converter.ConvertTo(target, typeof(byte[])) as byte[];
                return bytes;
            }
        }

        [TestMethod]
        public void TestMethod8()
        {
            var notepad = @"C:\Windows\System32\notepad.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", notepad);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var eyes = new Eyes();
            eyes.ApiKey = "lFMnO9e1nbFQEdhebrDmhEBOqB7jXUXJGogj103lRWAzM110";
            eyes.SetAppEnvironment("Windows 10", null);

            try
            {
                eyes.Open("AppiumPractice", "TestMethod8");
                Thread.Sleep(3000);
                var currentAppScreenshot = driver.GetScreenshot();
                var croppedScreenshot = CropScreenshot(currentAppScreenshot);

                eyes.CheckImage(croppedScreenshot, "Notepad Just Opened");

                eyes.Close();
            }
            finally
            {
                eyes?.AbortIfNotClosed();
                driver.Quit();
            }
        }

        [TestMethod]
        public void TestMethod9()
        {
            var notepad = @"C:\Windows\System32\notepad.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", notepad);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var textEditor = driver.FindElementByAccessibilityId("15");
            textEditor.SendKeys("ABCDEF");

            var eyes = new Eyes();
            eyes.ApiKey = "lFMnO9e1nbFQEdhebrDmhEBOqB7jXUXJGogj103lRWAzM110";
            eyes.SetAppEnvironment("Windows 10", null);

            try
            {
                eyes.Open("AppiumPractice", "TestMethod9");
                Thread.Sleep(3000);
                var currentAppScreenshot = driver.GetScreenshot();
                var croppedScreenshot = CropScreenshotWithRect(currentAppScreenshot, textEditor);

                eyes.CheckImage(croppedScreenshot, "Cropped Text Editor");

                eyes.Close();
            }
            finally
            {
                eyes?.AbortIfNotClosed();
                driver.Quit();
            }
        }

        [TestMethod]
        public void TestMethod10()
        {
            var notepad = @"C:\Windows\System32\notepad.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", notepad);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var textEditor = driver.FindElementByAccessibilityId("15");

            var wait = new OpenQA.Selenium.Support.UI.WebDriverWait(textEditor.WrappedDriver, TimeSpan.FromSeconds(30));
            textEditor.SendKeys("ABCDEF");
            wait.Until(ExpectedConditions.ElementToBeClickable(textEditor));
            textEditor.Click();
            textEditor.Click();
            textEditor.Click();
            textEditor.Click();
            textEditor.Click();
            textEditor.Click();

        }


        [TestMethod]
        public void TestMethod11()
        {
            var notepad = @"C:\Windows\System32\notepad.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", notepad);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);
            driver.Manage().Window.Maximize();

            var textEditor = driver.FindElementByAccessibilityId("15");
            textEditor.SendKeys("ABCDEF");

            var eyes = new Eyes();
            eyes.ApiKey = "lFMnO9e1nbFQEdhebrDmhEBOqB7jXUXJGogj103lRWAzM110";
            eyes.SetAppEnvironment("Windows 10", null);

            try
            {
                eyes.Open("AppiumPractice", "TestMethod11");
                Thread.Sleep(3000);
                var currentAppScreenshot = driver.GetScreenshot();
                var croppedScreenshot = CropScreenshotWithRect(currentAppScreenshot, textEditor);
                var croppedScreenshot2 = CropScreenshotOf2(croppedScreenshot, new Rectangle(900, 0, textEditor.Size.Width - 900, textEditor.Size.Height));

                eyes.CheckImage(croppedScreenshot2, "Cropped Text Editor");

                eyes.Close();
            }
            finally
            {
                eyes?.AbortIfNotClosed();
                driver.Quit();
            }
        }

        [TestMethod]
        public void TestMethod12()
        {
            var processes = Process.GetProcessesByName("M4Solutions");

            foreach(var process in processes)
            {
                process.Kill();
            }
        }

        [TestMethod]
        public void TestMethod13()
        {
            var paintbrush = @"C:\Windows\System32\mspaint.exe";

            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", paintbrush);

            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            var editor = driver.FindElementByAccessibilityId("59648");

            var action = new Actions(driver);
            action.MoveToElement(editor, 50, 50);
            action.Perform();

            action.Click();
            action.Perform();

            Thread.Sleep(5000);
        }

        [TestMethod]
        public void TestMethod14()
        {
            var options = new AppiumOptions();
            options.AddAdditionalCapability("app", @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
            
            var driver = new WindowsDriver<WindowsElement>(new Uri(" http://127.0.0.1:4723/"), options);

            Thread.Sleep(2000);

            var editBox = driver.FindElementByTagName("Edit");
            editBox.SendKeys(@"file:///C:\Users\cguzman\OneDrive - Resolvit\Desktop\pepe.pdf");
            editBox.SendKeys(Keys.Enter);

            Thread.Sleep(3000);

            driver.Quit();
        }
    }
}
