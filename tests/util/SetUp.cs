using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;

namespace SiggaPS.tests.util
{
    public class SetUp
    {
        public static IWebDriver driver;
        public static IWebDriver Driver
        {
            get
            {
                return driver;
            }

            set
            {
                driver = value;
            }
        }
        public SetUp()
        {

        }
        public static void TearDown()
        {
            try
            {
                driver.Close();
                driver.Quit();
            }
            catch (Exception) { }
        }
    }
}
