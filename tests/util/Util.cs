using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SiggaPS.tests.util
{
    class Util
    {
        public void RefreshPage(IWebDriver driver)
        {
            try
            {
                driver.Navigate().Refresh();
            }
            catch (Exception)
            {

            }
        }
        public void OpenPage(IWebDriver driver, string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);
            }
            catch (Exception) { }
        }
        public void ClickJS(IWebElement element, bool higth = true)
        {
            try
            {
                IJavaScriptExecutor executor = (IJavaScriptExecutor)SetUp.Driver;
                executor.ExecuteScript("arguments[0].click();", element);
            }
            catch (Exception) { }
        }
        public void ScrollElement(IWebDriver driver, IWebElement element)
        {
            try
            {
                if (element != null)
                {
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    js.ExecuteScript("arguments[0].scrollIntoView();", element);
                }
            }
            catch (Exception) { }
        }
        public string Screenshot(string strLog)
        {
                string filePath = System.IO.Directory.GetParent(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)).FullName).FullName).FullName + @"\SiggaPS\tests\imagesReport\" + strLog + ".PNG";
                ((ITakesScreenshot)SetUp.Driver).GetScreenshot().SaveAsFile(filePath, ScreenshotImageFormat.Png);
                return filePath;           
        }
        public void HighlightElementPassou(IWebElement elemento)
        {
            try
            {
                if (elemento != null)
                {
                   IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)SetUp.Driver;
                    javaScriptExecutor.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);",
                            elemento, "outline: 4px solid #00FF00;");
                }
            }
            catch (Exception) { }
        }
        public void HighlightElementFalhou(IWebElement elemento)
        {
            try
            {

                if (elemento == null || elemento != null)
                {
                   IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)SetUp.Driver;
                    javaScriptExecutor.ExecuteScript("arguments[0].setAttribute('style', arguments[1]);",
                            elemento, "outline: 4px solid #FF0000;");
                }
            }
            catch (Exception) { }
        }
                
    }
}
