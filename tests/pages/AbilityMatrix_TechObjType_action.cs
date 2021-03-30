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
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.pages
{
    class AbilityMatrix_TechObjType_action
    {
        public ExcelWorksheet xlsxInput { get; set; }
        public string excelAsset { get; set; }
        public string excelDescription { get; set; }
        public string excelQualification { get; set; }
        public string captureErrorAssetClass { get; set; }
        public string captureErrorDescription { get; set; }
        public string getPerson { get; set; }

        //construtor
        public AbilityMatrix_TechObjType_action()
        {
            try
            {
                PageFactory.InitElements(SetUp.Driver, this);
            }
            catch (Exception) { }
        }
        //atributos

        [FindsBy(How = How.XPath, Using = "//div[@id='__xmlview0--iconTabFilterTechnicalObjectType']")]
        [CacheLookup]
        public IWebElement buttonTechObjType { get; set; }

        [FindsBy(How = How.Id, Using = "createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-rows-row4-col1")]
        [CacheLookup]
        public IWebElement sapUiTableCellInner { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--buttonSave")]
        [CacheLookup]
        public IWebElement buttonSave { get; set; }

        [FindsBy(How = How.ClassName, Using = "sapMSLITitle")]
        [CacheLookup]
        public IWebElement innerTextPerson { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__select2-__clone286-arrow")]
        [CacheLookup]
        public IWebElement selectFirstQualificationBox { get; set; }

        //Metodos
        public bool selectTabTechObjType()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='__xmlview0--iconTabFilterTechnicalObjectType']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonTechObjType));
            if (objectDisplay.Contains(false)) { return false; }
            buttonTechObjType.Click();
            return true;
        }

        public bool saveButton()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--buttonSave")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSave));
            if (objectDisplay.Contains(false)) { return false; }
            buttonSave.Click();
            return true;
        }

        public bool selectQualification()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--tableObjectType-rows-row0")));
            objectDisplay.Add(validacao.ValidaElemVisivel(selectFirstQualificationBox));
            if (objectDisplay.Contains(false)) { return false; }
            selectFirstQualificationBox.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiSimpleFixFlexFlexContent")));
            Random rdn = new Random();
            int numbers;
            numbers = rdn.Next(1, 6);
            var randomNumberQualification = string.Format("__item5-__select2-__clone286-{0}", numbers);
            var selectNumberQualification = SetUp.Driver.FindElements(By.Id(randomNumberQualification)).FirstOrDefault();
            selectNumberQualification.Click();
            return true;
        }

        public bool validateTechObjTypeClass()
        {
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Show Technical Object Types"))
            {
                new AbilityMatrix_SapConnect().tableLList();
                xlsxInput = AbilityMatrix_LListExcel.XlsxInput;
            }
            else
            {
                new AbilityMatrix_SapConnect().tableAmPerson(getPerson);
                xlsxInput = AbilityMatrix_AmPersonExcel.XlsxInput;
             }
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--idIconTabBarNoIcons-content")));
            var planilha = xlsxInput;
            var linha = 2;
            IList<IWebElement> rows = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableTr"));
            IList<IWebElement> rowArrowDown = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst"));
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 0; pos <= qtdRowSheet; pos++)
            {
                int posTable = pos > 23 ? 23 : pos;
                excelAsset = planilha.Cells[linha, 1]?.Value?.ToString().Replace(" ", "");
                excelDescription = planilha.Cells[linha, 2]?.Value?.ToString().Replace(" ", "");
                excelQualification = planilha.Cells[linha, 3]?.Value?.ToString().Replace(" ", "");
                if (string.IsNullOrEmpty(excelAsset)) { break; }
                var stringSplitEquipment = rows[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string assetclass = stringSplitEquipment[0].Replace(" ", "");
                string description = stringSplitEquipment[1].Replace(" ", "");
                string qualification = stringSplitEquipment[2].Substring(0, 1);
                if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Set a grade for an Technical Object Type"))
                {
                    if ((excelAsset.Equals(assetclass) && excelDescription.Equals(description)) && excelQualification.Equals(qualification))
                    {
                        if (posTable < 23 || pos == qtdRowSheet)
                            new Util().HighlightElementPassou(rows[posTable]);
                    }
                }
                else if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Show Technical Object Types"))
                {
                    if (excelAsset.Equals(assetclass) && excelDescription.Equals(description))
                    {
                        if (posTable < 23 || pos == qtdRowSheet)
                            new Util().HighlightElementPassou(rows[posTable]);
                    }
                }
                else
                {
                    new Util().HighlightElementFalhou(rows[posTable]);
                    captureErrorAssetClass = excelAsset;
                    captureErrorDescription = excelDescription;
                    return false;
                }
                linha++;
                if (posTable == 23)
                {
                    rowArrowDown[posTable].SendKeys(Keys.Down);
                    rowArrowDown = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst"));
                    rows = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableTr"));
                    continue;
                }
            }
            return true;
        }


    }
}
