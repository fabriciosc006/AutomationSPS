using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Reflection;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.util
{
    [Binding]
    public class Hook
    {
        private static ExtentTest featureName;
        private static ExtentTest scenario;
        private static AventStack.ExtentReports.ExtentReports extent;
        Util util = new Util();

        [BeforeTestRun]
        public static void inicialize()
        {
            try
            {
                string filePath = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)).FullName).FullName).FullName + @"\Relatorio\";
                var htmlReport = new ExtentHtmlReporter(filePath);
                htmlReport.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
                extent = new AventStack.ExtentReports.ExtentReports();
                extent.AttachReporter(htmlReport);
            }
            catch
            {
                throw;
            }
            finally
            {
              // invoke ?jenkins?
            }
        }

        [AfterTestRun]
        public static void TearDownReport()
        {

            extent.Flush();
        }

        [BeforeFeature]
        public static void BeforeFeature()
        {
            //createTest
            featureName = extent.CreateTest<Feature>(FeatureContext.Current.FeatureInfo.Title);

        }

        [AfterStep]
        public void insertStepsReport()
        {
            var step = ScenarioStepContext.Current.StepInfo.StepDefinitionType.ToString();

            if (ScenarioContext.Current.TestError == null)
            {
                if (step == "Given")
                {
                    scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Pass("PASS");
                }
                else if (step == "When")
                {
                    scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Pass("PASS"); ;
                }
                else if (step == "And")
                {
                    scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text).Pass("PASS"); ;
                }
                else if (step == "Then")
                {
                    scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Pass("PASS");
                }
            }
            else if (ScenarioContext.Current.TestError != null)
            {
                var status = TestContext.CurrentContext.Result.Outcome.Status;
                var errorMessage = TestContext.CurrentContext.Result.Message;

                string screenShotPath = util.Screenshot("error");

                if (step == "Given" && status == TestStatus.Failed)
                {
                    scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Fail("FAIL" + errorMessage, MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "Given" && status == TestStatus.Inconclusive)
                {
                    scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Fail("INCONCLUSIVE - page or component is unstable", MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "When" && status == TestStatus.Failed)
                {
                    scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Fail("FAIL" + errorMessage, MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "When" && status == TestStatus.Inconclusive)
                {
                    scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Fail("INCONCLUSIVE - page or component is unstable", MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "And" && status == TestStatus.Failed)
                {
                    scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text).Fail("FAIL" + errorMessage, MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "And" && status == TestStatus.Inconclusive)
                {
                    scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text).Fail("INCONCLUSIVE - page or component is unstable", MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "Then" && status == TestStatus.Failed)
                {
                    scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Fail("FAIL" + errorMessage, MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
                else if (step == "Then" && status == TestStatus.Inconclusive)
                {
                    scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Fail("INCONCLUSIVE - page or component is unstable", MediaEntityBuilder.CreateScreenCaptureFromBase64String(ConvertFilePathToBase64(screenShotPath)).Build());
                }
            }
        }
        private string ConvertFilePathToBase64(string filePath)
        {
            byte[] imageArray = File.ReadAllBytes(filePath);
            string base64ImageRepresentation = Convert.ToBase64String(imageArray);
            return base64ImageRepresentation;
        }

        [BeforeScenario]
        public void StartTest(FeatureInfo feature)
        {
            SelectBrowser("Chrome");
            scenario = featureName.CreateNode<Scenario>(ScenarioContext.Current.ScenarioInfo.Title);
        }

        [AfterScenario]
        public void TearDown()
        {
            try
            {
                {
                    SetUp.TearDown();
                }
            }
            catch (Exception) { }

        }

        internal void SelectBrowser(string browserType)
        {
            switch (browserType)
            {
                case "Chrome":
                    ChromeOptions options = new ChromeOptions();
                    //options.AddArgument("--headless");
                    SetUp.Driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);
                    SetUp.Driver.Manage().Window.Maximize();
                    break;

                case "Firefox":
                    SetUp.Driver = new OpenQA.Selenium.Firefox.FirefoxDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
                    SetUp.Driver.Manage().Window.Maximize();
                    break;

                case "InternetExplorer":
                    SetUp.Driver = new OpenQA.Selenium.IE.InternetExplorerDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
                    SetUp.Driver.Manage().Window.Maximize();
                    break;
                default:
                    break;

            }

        }
        enum BrowserType
        {
            Chrome,
            Firefox,
            IE
        }
    }
}


