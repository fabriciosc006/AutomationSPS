using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using SiggaPS.tests.dataBaseSAP.MachineDowntimeCalendar;
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
    class MachineDowntimeCalendar_action
    {
        public string extrairString { get; set; }
        public string excelPlant { get; set; }
        public string captureErrorPlant { get; set; }
        public string getPlant { get; set; }
        public string getCode { get; set; }
        public string getFunctionalLocation { get; set; }
        public string getEquipment { get; set; }
        public string getStartDate { get; set; }
        public string getFinishDate { get; set; }
        public string getDataLimit { get; set; }
        public string excelLListCodeItem { get; set; }
        public string excelLListDescricao { get; set; }
        public string excelFuncLocLocation { get; set; }
        public string excelFuncLocDescription { get; set; }
        public string captureErrorLListCodeItem { get; set; }
        public string captureErrorFuncLocLocation { get; set; }
        public string captureErrorLeqpEqunr { get; set; }
        public string captureErrorLListCodeItemDescription { get; set; }
        public string getMouth { get; set; }
        public string start { get; set; }


        public MachineDowntimeCalendar_action()
        {
            try
            {
                PageFactory.InitElements(SetUp.Driver, this);
            }
            catch (Exception) { }
        }

        [FindsBy(How = How.ClassName, Using = "sapMSltLabel")]
        [CacheLookup]
        public IWebElement comboBoxPlant { get; set; }

        [FindsBy(How = How.Id, Using = "__button11")]
        [CacheLookup]
        public IWebElement buttonWorkCenter { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMLIB.sapMLIB-CTX.sapMLIBShowSeparator.sapMLIBTypeInactive.sapMLIBActionable.sapMLIBHoverable.sapMLIBFocusable.sapMSLI.sapMSLIWithDescription")]
        [CacheLookup]
        public IWebElement scrollDownWorkCenter { get; set; }

        [FindsBy(How = How.Id, Using = "__dialog0-searchField-I")]
        [CacheLookup]
        public IWebElement textBoxSearchCodeCenter { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='__xmlview0--funcionalLocationId-inner']")]
        [CacheLookup]
        public IWebElement textBoxLocationFunctional { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='__xmlview0--equipmentId-inner']")]
        [CacheLookup]
        public IWebElement textBoxEquipment { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromSearchId-icon")]
        [CacheLookup]
        public IWebElement textBoxStartDate { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromId-icon")]
        [CacheLookup]
        public IWebElement textBoxDowntimeStartDate { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromSearchId-cal--Head-B1")]
        [CacheLookup]
        public IWebElement buttonMonthLeft { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromId-cal--Head-B1")]
        [CacheLookup]
        public IWebElement buttonDowntimeMonthLeft { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToSearchId-cal--Head-B1")]
        [CacheLookup]
        public IWebElement buttonMonthRight { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToId-cal--Head-B1")]
        [CacheLookup]
        public IWebElement buttonDowntimeMonthRight { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromSearchId-cal--Head-B2")]
        [CacheLookup]
        public IWebElement buttonYearLeft { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromId-cal--Head-B2")]
        [CacheLookup]
        public IWebElement buttonDowntimeYearLeft { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToSearchId-cal--Head-B2")]
        [CacheLookup]
        public IWebElement buttonYearRight { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToId-cal--Head-B2")]
        [CacheLookup]
        public IWebElement buttonDowntimeYearRight { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToSearchId-icon")]
        [CacheLookup]
        public IWebElement textBoxFinishDate { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodToId-icon")]
        [CacheLookup]
        public IWebElement textBoxDowntimeFinishDate { get; set; }

        [FindsBy(How = How.ClassName, Using = "sapMSLITitle")]
        [CacheLookup]
        public IWebElement boxSearchCodeCenterResult { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMInputBaseContentWrapper.sapMInputBaseDisabledWrapper")]
        [CacheLookup]
        public IWebElement textBoxWorkCenter { get; set; }

        [FindsBy(How = How.Id, Using = "__xmlview0--dataLimitId-inner")]
        [CacheLookup]
        public IWebElement textBoxDataLimite { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='__button5']")]
        [CacheLookup]
        public IWebElement buttonSearchInTheTop { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@id='__button13-content']")]
        [CacheLookup]
        public IWebElement scrollDown { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapUiTableTreeIcon.sapUiTableTreeIconNodeClosed")]
        [CacheLookup]
        public IWebElement locationFreeIcon { get; set; }

        [FindsBy(How = How.Id, Using = "__button14-BDI-content")]
        [CacheLookup]
        public IWebElement buttonEquipaments { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMInputBaseContentWrapper.sapMInputBaseContentWrapperState.sapMInputBaseContentWrapperError")]
        [CacheLookup]
        public IWebElement redFieldsPeriodError { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMMessageToast.sapUiSelectable.sapContrast.sapContrastPlus.sapUiUserSelectable")]
        [CacheLookup]
        public IWebElement toastError { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromSearchId-inner")]
        [CacheLookup]
        public IWebElement textBoxPeriodStartAlter { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--functionalLocationTreeTableId-rowsel0")]
        [CacheLookup]
        public IWebElement boxFunctionalLocation { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--iconTabFilterForCreation")]
        [CacheLookup]
        public IWebElement buttonDowntimeInformation { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--hourButtonId-inner")]
        [CacheLookup]
        public IWebElement buttonTimeHour { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--hourToId-inner")]
        [CacheLookup]
        public IWebElement texBoxTimeHour { get; set; }

        [FindsBy(How = How.ClassName, Using = "sapMSltLabel")]
        [CacheLookup]
        public IWebElement texBoxDowntimeType { get; set; }

        [FindsBy(How = How.Id, Using = "__button10")]
        [CacheLookup]
        public IWebElement buttonSave { get; set; }

        [FindsBy(How = How.Id, Using = "__button9")]
        [CacheLookup]
        public IWebElement buttonDelete { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMBtnBase.sapMBtn.sapMDialogBeginButton.sapMBarChild")]
        [CacheLookup]
        public IWebElement buttonDeleteConfirmation { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__button17")]
        [CacheLookup]
        public IWebElement buttonEdit { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editStoppedMachineCalendarDialogId--hourFromId-inner")]
        [CacheLookup]
        public IWebElement textBoxTimeFrom { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editStoppedMachineCalendarDialogId--stoppedType-label")]
        [CacheLookup]
        public IWebElement buttonDowntimeType { get; set; }

        [FindsBy(How = How.Id, Using = "__button18-inner")]
        [CacheLookup]
        public IWebElement buttonSaveEdit { get; set; }

        public bool acesso(string url)
        {
            new Util().OpenPage(SetUp.Driver, url);
            return true;
        }

        public bool openPlantComboBox()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapMSltLabel")));
            objectDisplay.Add(validacao.ValidaElemVisivel(comboBoxPlant));
            if (objectDisplay.Contains(false)) { return false; }
            comboBoxPlant.Click();
            Thread.Sleep(200);
            return true;
        }
        public bool validatePlantsValidate()
        {
            new MachineDowntimeCalendar_SapConnect().tablePlants();
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            IWebElement listBox = SetUp.Driver.FindElement(By.CssSelector(".sapMSelectList.sapMSltList-CTX"));
            var rows = listBox.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = MachineDowntimeCalendar_PlantExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            foreach (IWebElement row in rows)
            {
                string getPositionComboBox = row.GetAttribute("innerText").Split('-')[0].Replace(" ", "");
                excelPlant = planilha.Cells[linha, 2]?.Value?.ToString();
                if (getPositionComboBox.Equals(excelPlant))
                {
                    objectDisplay.Add(validacao.ValidaElemContainsText(row, excelPlant));
                }
                else
                {
                    objectDisplay.Add(validacao.ValidaElemEqualsText(row, excelPlant));
                    captureErrorPlant = excelPlant;
                }
                if (objectDisplay.Contains(false)) { return false; }
                linha++;
            }
            return true;
        }
        public bool selectPLant(string plant)
        {
            getPlant = plant;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapMSltLabel")));
            objectDisplay.Add(validacao.ValidaElemVisivel(comboBoxPlant));
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(1000);
            comboBoxPlant.Click();
            Thread.Sleep(1000);
            IWebElement listBox = SetUp.Driver.FindElement(By.CssSelector(".sapMSelectList.sapMSltList-CTX"));
            var rows = listBox.FindElements(By.TagName("li"));
            foreach (IWebElement row in rows)
            {
                string getPositionComboBox = row.GetAttribute("innerText").Split('-')[0].Replace(" ", "");
                if (getPositionComboBox.Equals(plant))
                {
                    objectDisplay.Add(validacao.ValidaElemContainsText(row, plant));
                    row.Click();
                }
            }
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }
        public bool openWorkCenterButton()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonWorkCenter));
            if (objectDisplay.Contains(false)) { return false; }
            buttonWorkCenter.Click();
            Thread.Sleep(200);
            return true;
        }

        public bool validateWorkCentersRegistered()
        {
            new MachineDowntimeCalendar_SapConnect().tableLList(getPlant);
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMList.sapMListBGSolid")));
            IWebElement listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
            var rows = listWorkCenter.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = MachineDownTuneCalendar_LListExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 0; pos <= qtdRowSheet; pos++)
            {

                int posTable = pos > 5 ? 5 : pos;
                excelLListCodeItem = planilha.Cells[linha, 1]?.Value?.ToString();
                excelLListDescricao = planilha.Cells[linha, 2]?.Value?.ToString();
                if (string.IsNullOrEmpty(excelLListCodeItem)) { break; }
                var stringSplit = rows[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeItemScreen = extrairString = stringSplit[0];
                string descriptionScreen = extrairString = stringSplit[1];
                if (excelLListCodeItem.Equals(codeItemScreen) && excelLListDescricao.Equals(descriptionScreen))
                {
                    new Util().HighlightElementPassou(rows[pos]);
                    if (posTable < 5 || pos == qtdRowSheet) ;
                }
                else
                {
                    new Util().HighlightElementFalhou(rows[pos]);
                    captureErrorLListCodeItem = codeItemScreen;
                    captureErrorLListCodeItemDescription = descriptionScreen;
                    return false;
                }
                linha++;
                if (posTable == 5)
                {
                    rows[pos].SendKeys(Keys.Down);
                    listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
                    rows = listWorkCenter.FindElements(By.TagName("li"));
                    continue;
                }
            }
            return true;
        }
        public bool textBoxWorkCenterSearch(string code)
        {
            getCode = code;
            new MachineDowntimeCalendar_SapConnect().tableLListScenario3(getPlant, getCode);
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__dialog0-searchField-I")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxSearchCodeCenter));
            if (objectDisplay.Contains(false)) { return false; }
            textBoxSearchCodeCenter.SendKeys(code);
            Thread.Sleep(300);
            textBoxSearchCodeCenter.SendKeys(Keys.Enter);
            Thread.Sleep(1200);
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Filtering by plant and work center"))
            {
                if (boxSearchCodeCenterResult.GetAttribute("innerText").Equals(code))
                    boxSearchCodeCenterResult.Click();
            }
            return true;
        }

        public bool textBoxFunctionalLocationSearch(string code)
        {
            if (FeatureContext.Current.FeatureInfo.Title.Equals("MDC - Manage MDC"))
            {
                DateTime datetime = DateTime.Now;
                string datetimeValor = datetime.ToString("yyyyMMdd");
                new MachineDowntimeCalendar_SapConnect().cleanTableMachCal(datetimeValor);
            }
            getFunctionalLocation = code;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@id='__xmlview0--funcionalLocationId-inner']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxLocationFunctional));
            if (objectDisplay.Contains(false)) { return false; }
            textBoxLocationFunctional.SendKeys(code);
            return true;
        }

        public bool textBoxEquipmentSearch(string code)
        {
            getEquipment = code;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@id='__xmlview0--equipmentId-inner']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxEquipment));
            if (objectDisplay.Contains(false)) { return false; }
            textBoxEquipment.SendKeys(code);
            return true;
        }

        public bool textBoxStartDateFinishDate()
        {
            textBoxPeriodLeft();
            textBoxPeriodRight();
            return true;
        }

        public bool textBoxPeriodLeft()
        {

            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Filtering by plant and date"))
            {
                DateTime days = DateTime.Now.AddYears(-1);
                start = days.ToString("dd/MM/yyyy");
            }
            else
            {
                DateTime days = DateTime.Now.AddDays(+7);
                start = days.ToString("dd/MM/yyyy");
            }

            convertCalentarMonthPopPup(start);
            string dayStart = start.Substring(0, 2);
            string monthStart = start.Substring(3, 2);
            string yearStart = start.Substring(6, 4);
            var dateStart = DateTime.ParseExact(start, "dd/MM/yyyy", null);
            getStartDate = dateStart.ToString("yyyy/MM/dd").Replace("/", "");
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromSearchId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxStartDate));
            Thread.Sleep(200);
            textBoxStartDate.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromSearchId-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonMonthLeft));
            buttonMonthLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("#__xmlview0--periodFromSearchId-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromSearchId-cal--Head-B2")));
            buttonYearLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("#__xmlview0--periodFromSearchId-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearStart))
                {
                    yearSelect.Click();
                    break;
                }
            }
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxStartDate).DoubleClick().Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthView")));
            string firstDay = dayStart.Substring(0, 1);
            string secondDay = dayStart.Substring(1, 1);
            string firstMonth = monthStart.Substring(0, 1);
            string secondMonth = monthStart.Substring(1, 1);
            if (dayStart.StartsWith("0")) { dayStart = dayStart.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {

                var day = string.Format("#__xmlview0--periodFromSearchId-cal--Month0-2021{0}{1}{2}{3}", firstMonth, secondMonth, firstDay, secondDay);
                var daySelect = SetUp.Driver.FindElement(By.CssSelector(day));
                if (daySelect.GetAttribute("innerText").Equals(dayStart))
                {
                    daySelect.Click();
                    break;
                }
            }
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Filtering by plant and date"))
            {
                textBoxPeriodStartAlter.Clear();
                string send = dayStart + "/" + monthStart + "/" + yearStart;
                textBoxPeriodStartAlter.SendKeys(send);
            }

            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool textBoxPeriodRight()
        {
            DateTime now = DateTime.Now;
            string end = now.ToString("dd/MM/yyyy");
            convertCalentarMonthPopPup(end);
            string dayFinish = end.Substring(0, 2);
            string yearFinish = end.Substring(6, 4);
            var dateFinish = DateTime.ParseExact(end, "dd/MM/yyyy", null);
            getFinishDate = dateFinish.ToString("yyyy/MM/dd").Replace("/", "");
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToSearchId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxFinishDate));
            Thread.Sleep(200);
            textBoxFinishDate.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToSearchId-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonMonthRight));
            buttonMonthRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("#__xmlview0--periodToSearchId-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToSearchId-cal--Head-B2")));
            buttonYearRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("#__xmlview0--periodToSearchId-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearFinish))
                {
                    yearSelect.Click();
                    break;
                }
            }
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxFinishDate).DoubleClick().Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToSearchId-cal--Month0-WH6")));
            string first = dayFinish.Substring(0, 1);
            string second = dayFinish.Substring(1, 1);
            if (dayFinish.StartsWith("0")) { dayFinish = dayFinish.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {
                var day = string.Format("#__xmlview0--periodToSearchId-cal--Month0-202103{0}{1}", first, second);
                var daySelect = SetUp.Driver.FindElement(By.CssSelector(day));
                if (daySelect.GetAttribute("innerText").Equals(dayFinish))
                {
                    daySelect.Click();
                    break;
                }
            }

            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool textBoxDowntimeFututeStartDateFinishDate()
        {
            textBoxDowntimePeriodLeft();
            textBoxDowntimePeriodRight();
            return true;
        }

        public bool textBoxDowntimePeriodLeft()
        {
            DateTime days = DateTime.Now.AddDays(+2);
            start = days.ToString("dd/MM/yyyy");
            convertCalentarMonthPopPup(start);
            string dayStart = start.Substring(0, 2);
            string monthStart = start.Substring(3, 2);
            string yearStart = start.Substring(6, 4);
            var dateStart = DateTime.ParseExact(start, "dd/MM/yyyy", null);
            getStartDate = dateStart.ToString("yyyy/MM/dd").Replace("/", "");
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxDowntimeStartDate));
            Thread.Sleep(700);
            textBoxDowntimeStartDate.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromId-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDowntimeMonthLeft));
            buttonDowntimeMonthLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("#__xmlview0--periodFromId-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodFromId-cal--Head-B2")));
            buttonDowntimeYearLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("#__xmlview0--periodFromId-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearStart))
                {
                    yearSelect.Click();
                    break;
                }
            }
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxDowntimeStartDate).DoubleClick().Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthView")));
            string firstDay = dayStart.Substring(0, 1);
            string secondDay = dayStart.Substring(1, 1);
            string firstMonth = monthStart.Substring(0, 1);
            string secondMonth = monthStart.Substring(1, 1);
            if (dayStart.StartsWith("0")) { dayStart = dayStart.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {

                var day = string.Format("#__xmlview0--periodFromId-cal--Month0-2021{0}{1}{2}{3}", firstMonth, secondMonth, firstDay, secondDay);
                var daySelect = SetUp.Driver.FindElement(By.CssSelector(day));
                if (daySelect.GetAttribute("innerText").Equals(dayStart))
                {
                    daySelect.Click();
                    break;
                }
            }
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool textBoxDowntimePeriodRight()
        {
            DateTime now = DateTime.Now.AddDays(+7);
            string end = now.ToString("dd/MM/yyyy");
            convertCalentarMonthPopPup(end);
            string dayFinish = end.Substring(0, 2);
            string yearFinish = end.Substring(6, 4);
            var dateFinish = DateTime.ParseExact(end, "dd/MM/yyyy", null);
            getFinishDate = dateFinish.ToString("yyyy/MM/dd").Replace("/", "");
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxDowntimeFinishDate));
            Thread.Sleep(200);
            textBoxDowntimeFinishDate.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToId-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDowntimeMonthRight));
            buttonDowntimeMonthRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("#__xmlview0--periodToId-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToId-cal--Head-B2")));
            buttonDowntimeYearRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("#__xmlview0--periodToId-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearFinish))
                {
                    yearSelect.Click();
                    break;
                }
            }
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxDowntimeFinishDate).DoubleClick().Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--periodToId-cal--Month0-WH6")));
            string first = dayFinish.Substring(0, 1);
            string second = dayFinish.Substring(1, 1);
            if (dayFinish.StartsWith("0")) { dayFinish = dayFinish.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {
                var day = string.Format("#__xmlview0--periodToId-cal--Month0-202103{0}{1}", first, second);
                var daySelect = SetUp.Driver.FindElement(By.CssSelector(day));
                if (daySelect.GetAttribute("innerText").Equals(dayFinish))
                {
                    daySelect.Click();
                    break;
                }
            }
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool buttonFillTime()
        {            
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--hourButtonId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonTimeHour));
            buttonTimeHour.Click();
            Thread.Sleep(200);
            if(texBoxTimeHour.GetAttribute("value") != null)
            {
                objectDisplay.Add(validacao.ValidaElemVisivel(texBoxTimeHour));
            }
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool selectDowntimeType()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapMSltLabel")));
            objectDisplay.Add(validacao.ValidaElemVisivel(texBoxDowntimeType));
            texBoxDowntimeType.Click();
            Thread.Sleep(200);
            IList<IWebElement> listType = SetUp.Driver.FindElements(By.CssSelector(".sapMSelectList.sapMSltList-CTX"));
            listType[1].Click();
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool saveDowntimeInformation()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__button10")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSave));
            if (objectDisplay.Contains(false)) { return false; }
            buttonSave.Click();           
            return true;
        }

        public bool deleteDowntimeInformation()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__button9")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDelete));
            if (objectDisplay.Contains(false)) { return false; }
            buttonDelete.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMBtnBase.sapMBtn.sapMDialogBeginButton.sapMBarChild")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDeleteConfirmation));
            buttonDeleteConfirmation.Click();
            Thread.Sleep(2000);
            return true;
        }

        public bool buttonEditMachine()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__button17")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonEdit));
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(700);
            buttonEdit.Click();
            return true;
        }

        public bool editTimeFrom()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#editStoppedMachineCalendarDialogId--hourFromId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxTimeFrom));
            textBoxTimeFrom.SendKeys("0707");
            Thread.Sleep(200);
            buttonDowntimeType.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("editStoppedMachineCalendarDialogId--stoppedType-valueStateText")));
            IList<IWebElement> list = SetUp.Driver.FindElements(By.Id("editStoppedMachineCalendarDialogId--stoppedType-valueStateText"));
            list[0].Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__button18-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSaveEdit));
            buttonSaveEdit.Click();
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public void convertCalentarMonthPopPup(string month)
        {
            month = month.Substring(3, 2);

            Dictionary<string, string> monthDictionary = new Dictionary<string, string>()
         {
                { "01","January"},
                { "02","February"},
                { "03","March"},
                { "04","April"},
                { "05","May"},
                { "06","June" },
                { "07","July"},
                { "08","August"},
                { "09","September"},
                { "10","September"},
                { "11","November"},
                { "12","December"},
         };
            foreach (var find in monthDictionary)
            {
                if (month.Contains(find.Key))
                {
                    getMouth = find.Value;
                    break;
                }
            }
        }
        public bool validateWorkCentersRegisteredScenario3()
        {
            new MachineDowntimeCalendar_SapConnect().tableLListScenario3(getPlant, getCode);
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMList.sapMListBGSolid")));
            IWebElement listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
            var rows = listWorkCenter.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = MachineDownTuneCalendar_LListExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 0; pos <= qtdRowSheet; pos++)
            {

                int posTable = pos > 5 ? 5 : pos;
                excelLListCodeItem = planilha.Cells[linha, 1]?.Value?.ToString();
                excelLListDescricao = planilha.Cells[linha, 2]?.Value?.ToString();
                if (string.IsNullOrEmpty(excelLListCodeItem)) { break; }
                var stringSplit = rows[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeItemScreen = extrairString = stringSplit[0];
                string descriptionScreen = extrairString = stringSplit[1];
                if (excelLListCodeItem.Equals(codeItemScreen) && excelLListDescricao.Equals(descriptionScreen))
                {
                    new Util().HighlightElementPassou(rows[pos]);
                    if (posTable < 5 || pos == qtdRowSheet) ;
                }
                else
                {
                    new Util().HighlightElementFalhou(rows[pos]);
                    captureErrorLListCodeItem = codeItemScreen;
                    captureErrorLListCodeItemDescription = descriptionScreen;
                    return false;
                }
                linha++;
                if (posTable == 5)
                {
                    rows[pos].SendKeys(Keys.Down);
                    listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
                    rows = listWorkCenter.FindElements(By.TagName("li"));
                    continue;
                }
            }
            return true;
        }
        public bool deleteWorkCenter()
        {
            new MachineDowntimeCalendar_SapConnect().tableLListScenario3(getPlant, getCode);
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMList.sapMListBGSolid")));
            IWebElement listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
            var rows = listWorkCenter.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = MachineDownTuneCalendar_LListExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 0; pos <= qtdRowSheet; pos++)
            {
                excelLListCodeItem = planilha.Cells[linha, 1]?.Value?.ToString();
                excelLListDescricao = planilha.Cells[linha, 2]?.Value?.ToString();
                if (string.IsNullOrEmpty(excelLListCodeItem)) { break; }
                var stringSplit = rows[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeItemScreen = extrairString = stringSplit[0];
                string descriptionScreen = extrairString = stringSplit[1];
                if (excelLListCodeItem.Equals(codeItemScreen) && excelLListDescricao.Equals(descriptionScreen))
                {
                    new Util().HighlightElementPassou(rows[pos]);
                    rows[pos].Click();
                    break;
                    Thread.Sleep(500);
                }
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapUiSmallMarginBottom")));
            IList<IWebElement> buttonClear = SetUp.Driver.FindElements(By.CssSelector(".sapUiSmallMarginBottom"));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonClear[1]));
            if (objectDisplay.Contains(false)) { return false; }
            buttonClear[1].Click();
            return true;
        }
        public bool validateButtonClearAndTextBox()
        {
            if (string.IsNullOrEmpty(textBoxWorkCenter.GetAttribute("innerText")))
            {
                new Util().HighlightElementPassou(textBoxWorkCenter);
            }
            else
            {
                new Util().HighlightElementFalhou(textBoxWorkCenter);
            }

            return true;
        }
        public bool changeDataLimit(string datalimit)
        {
            getDataLimit = datalimit;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__xmlview0--dataLimitId-inner")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxDataLimite));
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(1000);
            textBoxDataLimite.Clear();
            Thread.Sleep(1000);
            textBoxDataLimite.Click();
            Thread.Sleep(1000);
            textBoxDataLimite.SendKeys(datalimit);
            Thread.Sleep(1000);
            return true;
        }
        public bool buttonSearch()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[@id='__button5']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSearchInTheTop));
            if (objectDisplay.Contains(false)) { return false; }
            buttonSearchInTheTop.Click();
            Thread.Sleep(1000);
            if (!ScenarioContext.Current.ScenarioInfo.Title.Equals("Filtering by plant and an invalid date"))
            {
                ((IJavaScriptExecutor)SetUp.Driver).ExecuteScript("arguments[0].scrollIntoView(true);", scrollDown);
            }
            return true;
        }

        public bool fieldsPeriodWithRed()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMInputBaseContentWrapper.sapMInputBaseContentWrapperState.sapMInputBaseContentWrapperError")));
            IList<IWebElement> periodsReds = SetUp.Driver.FindElements(By.CssSelector(".sapMInputBaseContentWrapper.sapMInputBaseContentWrapperState.sapMInputBaseContentWrapperError"));
            objectDisplay.Add(validacao.ValidaElemVisivel(periodsReds[0]));
            objectDisplay.Add(validacao.ValidaElemVisivel(periodsReds[1]));
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool toastErrorMessage()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMMessageToast.sapUiSelectable.sapContrast.sapContrastPlus.sapUiUserSelectable")));
            objectDisplay.Add(validacao.ValidaElemVisivel(toastError));
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool selectFunctional()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--functionalLocationTreeTableId-rowsel0")));
            IList<IWebElement> boxList = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableRowSelectionCell"));
            boxList[0].Click();
            if(ScenarioContext.Current.ScenarioInfo.Title.Equals("Registering a downtime for two (or more) functional locations"))
            {
                boxList[2].Click();
                boxList[3].Click();
            }
            textBoxDataLimite.SendKeys(Keys.PageUp);
            return true;
        }

        public bool selectTabEquipments()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("__button14-BDI-content")));
            buttonEquipaments.Click();           
            return true;
        }

        public bool selectEquipments()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableRowSelectionCell")));
            IList<IWebElement> boxList = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableRowSelectionCell"));
            boxList[0].Click();
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Registering a downtime for two (or more) functional locations")
                || ScenarioContext.Current.ScenarioInfo.Title.Equals("Registering a downtime for two (or more) equipments"))
            {
                boxList[2].Click();
                boxList[3].Click();
            }
            textBoxDataLimite.SendKeys(Keys.PageUp);
            return true;
        }

        public bool selectDataMachineCalendar()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            IWebElement boxOne = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#__xmlview0--machineCalendarTableId-rowsel0")));
            boxOne.Click();
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Deleting two (or more) downtimes"))
            {
                IWebElement boxTwo = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#__xmlview0--machineCalendarTableId-rowsel1")));
                boxTwo.Click();
            }
            textBoxDataLimite.SendKeys(Keys.PageUp);
            return true;
        }

        public bool clickTabDowntimeInformation()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--iconTabFilterForCreation")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDowntimeInformation));
            if (objectDisplay.Contains(false)) { return false; }
            buttonDowntimeInformation.Click();
            return true;
        }

        public bool validateLocations()
        {
            new MachineDowntimeCalendar_SapConnect().tableFuncLoc(getPlant, getDataLimit, getCode, getFunctionalLocation);
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@id='__xmlview0--functionalLocationTreeTableId-table']")));
            IList<IWebElement> rowsLocation = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst.sapUiTableCellFlex"));
            IList<IWebElement> rowsLocationDescription = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellLast"));
            ExcelWorksheet xlsxInputLocation = MachineDowntimeCalendar_FuncLocExcel.XlsxInput;
            var planilhaLocation = xlsxInputLocation;
            var linha = 2;
            var qtdRowSheet = planilhaLocation.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 0; pos <= qtdRowSheet; pos++)
            {
                int posTable = pos > 6 ? 6 : pos;
                var icon = string.Format("#__xmlview0--functionalLocationTreeTableId-rows-row{0}-treeicon", posTable);
                var locationIconExpand = SetUp.Driver.FindElements(By.CssSelector(icon)).FirstOrDefault();
                excelFuncLocLocation = planilhaLocation.Cells[linha, 1]?.Value?.ToString().Replace(" ", "");
                excelFuncLocDescription = planilhaLocation.Cells[linha, 2]?.Value?.ToString().Replace(" ", "");
                if (string.IsNullOrEmpty(excelFuncLocLocation)) { break; }
                var stringSplitLocation = rowsLocation[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeFuncLocationScreen = extrairString = stringSplitLocation[0].Replace(" ", "");
                var stringSplitDescription = rowsLocationDescription[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeDescriptionScreen = extrairString = stringSplitDescription[0].Replace(" ", "");
                if ((locationIconExpand != null && !locationIconExpand.GetAttribute("aria-hidden").Equals("true")) && (excelFuncLocLocation.Equals(codeFuncLocationScreen) && excelFuncLocDescription.Equals(codeDescriptionScreen)))
                {
                    if (posTable < 6 || pos == qtdRowSheet)
                        new Util().HighlightElementPassou(rowsLocation[posTable]);
                    new Util().HighlightElementPassou(rowsLocationDescription[posTable]);
                    locationIconExpand?.Click();
                    Thread.Sleep(200);
                }
                else if (excelFuncLocLocation.Equals(codeFuncLocationScreen) && excelFuncLocDescription.Equals(codeDescriptionScreen))
                {
                    if (posTable < 6 || pos == qtdRowSheet)
                        new Util().HighlightElementPassou(rowsLocation[posTable]);
                    new Util().HighlightElementPassou(rowsLocationDescription[posTable]);
                }
                else
                {
                    captureErrorFuncLocLocation = excelFuncLocLocation;
                    new Util().HighlightElementFalhou(rowsLocation[posTable]);
                    new Util().HighlightElementFalhou(rowsLocationDescription[posTable]);
                    return false;
                }
                linha++;
                if (posTable == 6)
                {
                    rowsLocation[posTable].SendKeys(Keys.Down);
                    rowsLocation = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst.sapUiTableCellFlex"));
                    rowsLocationDescription = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellLast"));
                    continue;
                }
            }
            return true;
        }
        public bool validateEquipments()
        {
            new MachineDowntimeCalendar_SapConnect().tableLEQP(getPlant, getDataLimit, getCode, getEquipment);
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("__button14-BDI-content")));
            buttonEquipaments.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@id='__xmlview0--equipmentTreeTableId-table']")));
            IList<IWebElement> rowsEquipment = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst.sapUiTableCellFlex"));
            IList<IWebElement> rowsEquipmentDescription = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellLast"));
            ExcelWorksheet xlsxInput = MachineDowntimeCalendar_LeqpExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            List<MachineDowntimeCalendar_CellsDTO> sheetObject = new List<MachineDowntimeCalendar_CellsDTO>();
            for (int x = 1; x <= qtdRowSheet; x++)
            {
                sheetObject.Add(new MachineDowntimeCalendar_CellsDTO
                {
                    Number = planilha.Cells[x, 1]?.Value?.ToString(),
                    Text = planilha.Cells[x, 2]?.Value?.ToString().Replace(" ", "")
                });
            }
            for (int pos = 0; pos < qtdRowSheet; pos++)
            {
                int posTable = pos > 6 ? 6 : pos;
                var icon = string.Format("#__xmlview0--equipmentTreeTableId-rows-row{0}-treeicon", posTable);
                var locationIconExpand = SetUp.Driver.FindElements(By.CssSelector(icon)).FirstOrDefault();
                var stringSplitEquipment = rowsEquipment[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeEquipmentScreen = extrairString = stringSplitEquipment[0];
                var stringEquipmentDescription = rowsEquipmentDescription[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeDescriptionScreen = extrairString = stringEquipmentDescription[0].Replace(" ", "");
                if ((locationIconExpand?.GetAttribute("aria-hidden") != null && !locationIconExpand.GetAttribute("aria-hidden").Equals("true"))
                    && ((sheetObject
                    .Where(x => x.Number.Equals(codeEquipmentScreen)
                        && x.Text.Equals(codeDescriptionScreen))
                    .FirstOrDefault() != null)))
                {
                    if (posTable < 6 || pos == qtdRowSheet)
                        new Util().HighlightElementPassou(rowsEquipment[posTable]);
                    new Util().HighlightElementPassou(rowsEquipmentDescription[posTable]);
                    locationIconExpand?.Click();
                    Thread.Sleep(200);
                }
                else if (sheetObject
                    .Where(x => x.Number.Equals(codeEquipmentScreen)
                        && x.Text.Equals(codeDescriptionScreen))
                    .FirstOrDefault() != null)

                {
                    var teste = locationIconExpand?.GetAttribute("aria-hidden");
                    if (posTable < 6 || pos == qtdRowSheet)
                        new Util().HighlightElementPassou(rowsEquipment[posTable]);
                    new Util().HighlightElementPassou(rowsEquipmentDescription[posTable]);
                }
                else
                {
                    captureErrorLeqpEqunr = codeDescriptionScreen;
                    new Util().HighlightElementFalhou(rowsEquipment[posTable]);
                    new Util().HighlightElementFalhou(rowsEquipmentDescription[posTable]);
                    return false;
                }
                linha++;
                if (posTable == 6)
                {
                    rowsEquipment[posTable].SendKeys(Keys.Down);
                    rowsEquipment = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellFirst.sapUiTableCellFlex"));
                    rowsEquipmentDescription = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableCell.sapUiTableContentCell.sapUiTableDataCell.sapUiTableCellLast"));
                    continue;
                }
            }
            return true;
        }
        public bool validateMachineCalendar()
        {
            if (FeatureContext.Current.FeatureInfo.Title.Equals("MDC - Manage MDC"))
            {
                new MachineDowntimeCalendar_SapConnect().tableMachCalFeatureManage(getPlant, getFunctionalLocation);
            }
            else
            {
                new MachineDowntimeCalendar_SapConnect().tableMachCal(getPlant, getCode, getFunctionalLocation, getEquipment, getStartDate, getFinishDate);
            }
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--machineCalendarTableId-table")));
            ExcelWorksheet xlsxInput = MachineDowntimeCalendar_Mach_CalExcel.XlsxInput;
            var planilha = xlsxInput;
            if (planilha is null)
            {
                IWebElement noData = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--machineCalendarTableId-noDataMsg"));
                new Util().HighlightElementFalhou(noData);
                return false;
            }
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            List<MachineDowntimeCalendar_CellsDTO> sheetObject = new List<MachineDowntimeCalendar_CellsDTO>();
            for (int x = 2; x <= qtdRowSheet; x++)
            {
                sheetObject.Add(new MachineDowntimeCalendar_CellsDTO
                {
                    Instalation = planilha.Cells[x, 1]?.Value?.ToString(),
                    DescriptionPltxt = planilha.Cells[x, 2]?.Value?.ToString().Replace(" ", ""),
                    Equipments = planilha.Cells[x, 3]?.Value?.ToString(),
                    DescriptionEqktx = planilha.Cells[x, 4]?.Value?.ToString().Replace(" ", ""),
                    InitialData = planilha.Cells[x, 5]?.Value?.ToString(),
                    finishData = planilha.Cells[x, 6]?.Value?.ToString(),
                    InitialTime = planilha.Cells[x, 7]?.Value?.ToString(),
                    finishTime = planilha.Cells[x, 8]?.Value?.ToString(),
                    paradeType = planilha.Cells[x, 9]?.Value?.ToString(),
                    paradeAlocation = planilha.Cells[x, 10]?.Value?.ToString(),
                    observation = planilha.Cells[x, 11]?.Value?.ToString().Replace(" ", ""),
                    index = x - 2
                });
            }
            int column = 11;
            for (int pos = 0; pos < qtdRowSheet - 1; pos++)
            {
                bool highlight = pos < 6 || pos == qtdRowSheet - 1;
                var sheetRows = sheetObject;
                for (int i = 0; i < column; i++)
                {
                    var columnCalendar = string.Format("#__xmlview0--machineCalendarTableId-rows-row{0}-col{1}",
                        pos > 6 ? 6 : pos, i);
                    var nextColumn = SetUp.Driver.FindElements(By.CssSelector(columnCalendar)).FirstOrDefault();
                    var nextColumnValue = nextColumn.GetAttribute("innerText")?.Replace(" ", "");

                    sheetRows = sheetRows
                                                .Where(row => nextColumnValue.Equals(sheetObjectValueByIndex(row, i)))
                                                .ToList();
                    //pintar tela
                    if (sheetRows is null || sheetRows.Count() == 0)
                    {
                        new Util().HighlightElementFalhou(nextColumn);
                        return false;
                    }
                    else if (highlight)
                    {
                        new Util().HighlightElementPassou(nextColumn);
                    }

                    ColumnRowsNavigate(nextColumn, pos, i);
                }
            }
            return true;
        }

        private string sheetObjectValueByIndex(MachineDowntimeCalendar_CellsDTO sheetObject, int columnIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return sheetObject.Instalation;
                case 1:
                    return sheetObject.DescriptionPltxt;
                case 2:
                    return sheetObject.Equipments;
                case 3:
                    return sheetObject.DescriptionEqktx;
                case 4:
                    return sheetObject.InitialData;
                case 5:
                    return sheetObject.finishData;
                case 6:
                    return sheetObject.InitialTime;
                case 7:
                    return sheetObject.finishTime;
                case 8:
                    return sheetObject.paradeType;
                case 9:
                    return sheetObject.paradeAlocation;
                case 10:
                    return sheetObject.observation;
                default:
                    return null;
            }
        }

        private void ColumnRowsNavigate(IWebElement nextColumn, int pos, int columnIndex)
        {
            if (columnIndex < 10)
            {
                nextColumn.SendKeys(Keys.ArrowRight);
            }
            else
            {
                IWebElement leftColumn = null;
                for (int j = 10; j >= 0; j--)
                {
                    var leftColumnCalendar = string.Format("#__xmlview0--machineCalendarTableId-rows-row{0}-col{1}",
                        pos > 6 ? 6 : pos, j);
                    leftColumn = SetUp.Driver.FindElements(By.CssSelector(leftColumnCalendar)).FirstOrDefault();
                    leftColumn.SendKeys(Keys.ArrowLeft);
                }
                leftColumn.SendKeys(Keys.ArrowDown);
            }
        }
    }
}


