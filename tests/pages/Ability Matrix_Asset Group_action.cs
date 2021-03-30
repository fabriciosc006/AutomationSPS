using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using SiggaPS.tests.dataBaseSAP.AbilityMatrix;
using SiggaPS.tests.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.pages
{
    class Ability_Matrix_Asset_Group_action
    {
        public string getCode { get; set; }

        //construtor
        public Ability_Matrix_Asset_Group_action()
        {
            try
            {
                PageFactory.InitElements(SetUp.Driver, this);
            }
            catch (Exception) { }
        }
        public bool acesso(string url)
        {
            new Util().OpenPage(SetUp.Driver, url);
            return true;
        }

        //atributos

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--iconTabFilterAssetGroup-text")]
        [CacheLookup]
        public IWebElement buttonAssetGroup { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--radionButtonAssetGroupEquipment-Button")]
        [CacheLookup]
        public IWebElement radioequipment { get; set; }
        
        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--buttonAssetGroupSelect-inner")]
        [CacheLookup]
        public IWebElement btnselect { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#selectDialogEquipment-searchField-I")]
        [CacheLookup]
        public IWebElement txbsearchitem { get; set; }

        [FindsBy(How = How.Id, Using = "selectDialogEquipment-list-listUl")]
        [CacheLookup]
        public IWebElement boxSearchEquipment { get; set; }

        //Metodos
        public bool selectTabAssetGroup()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='__xmlview0--iconTabFilterAssetGroup-text']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonAssetGroup));
            if (objectDisplay.Contains(false)) { return false; }
            buttonAssetGroup.Click();
            return true;
        }

        public bool gotogrouplist()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMRbBLabel")));
            objectDisplay.Add(validacao.ValidaElemVisivel(radioequipment));
            if (objectDisplay.Contains(false)) { return false; }

            if (ScenarioContext.Current.ScenarioInfo.Title.Contains("functional location"))
            {
                radioequipment.SendKeys(Keys.ArrowRight);
                Thread.Sleep(500);
                btnselect.Click();
            }
            else
            {
                btnselect.Click();
            }                 
         return true;
        }

        public bool textBoxAssetGroupSearch(string code)
        {
            getCode = code;          
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("selectDialogEquipment-dialog-header-BarPH")));
            objectDisplay.Add(validacao.ValidaElemVisivel(txbsearchitem));
            if (objectDisplay.Contains(false)) { return false; }
            txbsearchitem.SendKeys(code);
            Thread.Sleep(300);
            txbsearchitem.SendKeys(Keys.Enter);
            Thread.Sleep(1200);
            
            if (ScenarioStepContext.Current.StepInfo.Text.Equals("I search for an equipment")|| ScenarioStepContext.Current.StepInfo.Text.Equals("I search for a functional location"))
            {
                if (txbsearchitem.GetAttribute("value").Equals(code))
                    boxSearchEquipment.Click();
                    Thread.Sleep(500);
            }
         return true;
        }

        public bool validateAssetGroupEquipment()
       {
        //    if (FeatureContext.Current.FeatureInfo.Title.Equals("Ability Matrix - Asset Group"))
        //    {
        //        new AbilityMatrix_SapConnect().tableAssetGroupAmPerson(getCode);
        //    }
        //    else
        //    {
        //      //  new AbilityMatrix_SapConnect().tableMachCal(getPlant, getCode, getFunctionalLocation, getEquipment, getStartDate, getFinishDate);
        //    }
        //    WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--tableAssetGroup")));
        ////    ExcelWorksheet xlsxInput = MachineDowntimeCalendar_Mach_CalExcel.XlsxInput;
        //    var planilha = xlsxInput;
            //if (planilha is null)
            //{
            //    IWebElement noData = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--tableAssetGroup"));
            //    new Util().HighlightElementFalhou(noData);
            //    return false;
            //}
            //var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            //List<AbilityMatrix_AssetGroup_AmPerson_CellsDTO> sheetObject = new List<AbilityMatrix_AssetGroup_AmPerson_CellsDTO>();
            //for (int x = 2; x <= qtdRowSheet; x++)
            //{
            //    sheetObject.Add(new AbilityMatrix_AssetGroup_AmPerson_CellsDTO
            //    {
            //        AM_KEY_TYPE = planilha.Cells[x, 1]?.Value?.ToString(),
            //        AM_KEY = planilha.Cells[x, 2]?.Value?.ToString().Replace(" ", ""),
            //        EQKTX = planilha.Cells[x, 3]?.Value?.ToString(),
            //        AM_KN = planilha.Cells[x, 4]?.Value?.ToString().Replace(" ", ""),
            //                        
            //        index = x - 2
            //    });
            //}
            //int column = 7;
            //for (int pos = 0; pos < qtdRowSheet - 1; pos++)
            //{
            //    bool highlight = pos < 6 || pos == qtdRowSheet - 1;
            //    var sheetRows = sheetObject;
            //    for (int i = 0; i < column; i++)
            //    {
            //        var columnCalendar = string.Format("#__xmlview0--machineCalendarTableId-rows-row{0}-col{1}",
            //            pos > 6 ? 6 : pos, i);
            //        var nextColumn = SetUp.Driver.FindElements(By.CssSelector(columnCalendar)).FirstOrDefault();
            //        var nextColumnValue = nextColumn.GetAttribute("innerText")?.Replace(" ", "");

            //        sheetRows = sheetRows
            //                                    .Where(row => nextColumnValue.Equals(sheetObjectValueByIndex(row, i)))
            //                                    .ToList();
            //        //pintar tela
            //        if (sheetRows is null || sheetRows.Count() == 0)
            //        {
            //            new Util().HighlightElementFalhou(nextColumn);
            //            return false;
            //        }
            //        else if (highlight)
            //        {
            //            new Util().HighlightElementPassou(nextColumn);
            //        }

            //        ColumnRowsNavigate(nextColumn, pos, i);
            //    }
            //}

            return true;
        }


    }
}
