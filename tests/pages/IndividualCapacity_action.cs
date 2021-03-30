using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using SiggaPS.tests.dataBaseSAP.IndividualCapacity;
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
    class IndividualCapacity_action
    {
        ExcelWorksheet xlsxInputPeople { get; set; }
        public string excelPlant { get; set; }
        public string captureErrorPlant { get; set; }
        public string getPlant { get; set; }
        public string excelLListCodeItem { get; set; }
        public string excelLListDescricao { get; set; }
        public string captureErrorLListCodeItem { get; set; }
        public string captureErrorLListCodeItemDescription { get; set; }
        public string excelLPerson { get; set; }
        public string excelLPersonUsername { get; set; }
        public string captureErrorLPerson { get; set; }
        public string captureErrorLPersonUsername { get; set; }
        public string getCode { get; set; }
        public string getDataLimit { get; set; }
        public string getOnePerson { get; set; }
        public string getOnePersonCopyCapacity { get; set; }
        public string getMouth { get; set; }
        public string getStartDate { get; set; }
        public string getFinishDate { get; set; }
        public string start { get; set; }
        public string calendarStart { get; set; }
        public string calendarMiddle { get; set; }
        public string end { get; set; }
        public string getTimeStart { get; set; }
        public string getTimeFinish { get; set; }
        public DateTime days { get; set; }

        public List<string> getMorePerson = new List<string>().Distinct().ToList();

        public string personToString()
        {
            return String.Join(",", getMorePerson);
        }

        public IndividualCapacity_action()
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

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--personId-vhi")]
        [CacheLookup]
        public IWebElement buttonPeople { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--personId--valueHelpDialog-table-rows-row4-col0")]
        [CacheLookup]
        public IWebElement rowPeopleDown { get; set; }

        [FindsBy(How = How.Id, Using = "__dialog0-searchField-I")]
        [CacheLookup]
        public IWebElement textBoxSearchCodeCenter { get; set; }

        [FindsBy(How = How.Id, Using = "__xmlview0--dataLimitId-inner")]
        [CacheLookup]
        public IWebElement textBoxDataLimite { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='__button5']")]
        [CacheLookup]
        public IWebElement buttonSearchInTheTop { get; set; }

        [FindsBy(How = How.Id, Using = "__label9")]
        [CacheLookup]
        public IWebElement scrollDown { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--periodFromId-inner")]
        [CacheLookup]
        public IWebElement textBoxPeriodStartAlterPeriod { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--allocationFromId-inner")]
        [CacheLookup]
        public IWebElement textBoxPeriodStartAlterAlocation { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#cloneDialog--periodFromId-inner")]
        [CacheLookup]
        public IWebElement textBoxPeriodStartCloneCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--hourFromId-inner")]
        [CacheLookup]
        public IWebElement texBoxTimeHourStart { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__xmlview0--hourToId-inner")]
        [CacheLookup]
        public IWebElement texBoxTimeHourFinish { get; set; }

        [FindsBy(How = How.Id, Using = "__button9")]
        [CacheLookup]
        public IWebElement buttonDelete { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMBtnBase.sapMBtn.sapMDialogBeginButton.sapMBarChild")]
        [CacheLookup]
        public IWebElement buttonDeleteConfirmation { get; set; }

        [FindsBy(How = How.Id, Using = "__button10")]
        [CacheLookup]
        public IWebElement buttonSave { get; set; }

        [FindsBy(How = How.Id, Using = "__error0")]
        [CacheLookup]
        public IWebElement popUpExist { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__button13-__clone20-internalBtn")]
        [CacheLookup]
        public IWebElement buttonInternalFromTable { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[contains(text(),'Edit')]")]
        [CacheLookup]
        public IWebElement buttonInternalFromTableEdit { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[contains(text(),'Copy')]")]
        [CacheLookup]
        public IWebElement buttonInternalFromTableCopy { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editDialog--hourEditFromId-inner")]
        [CacheLookup]
        public IWebElement textBoxTimeStartEditCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editDialog--hourEditToId-inner")]
        [CacheLookup]
        public IWebElement textBoxTimeFinishEditCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editDialog--noteEditId-inner")]
        [CacheLookup]
        public IWebElement textBoxNoteEditCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#editDialog--editDialog-ok")]
        [CacheLookup]
        public IWebElement buttonOkEditCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#cloneDialog--cloneDialog-ok")]
        [CacheLookup]
        public IWebElement buttonOkCopyCapacity { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMSFI")]
        [CacheLookup]
        public IWebElement textBoxSearchCloneCapacity { get; set; }

        public bool acesso(string url)
        {
            new Util().OpenPage(SetUp.Driver, url);
            return true;
        }

        public bool clickComboBoxPlant()
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

        public bool openPeopleButton()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--personId-vhi")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonPeople));
            if (objectDisplay.Contains(false)) { return false; }
            buttonPeople.Click();
            Thread.Sleep(200);
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
            if (!ScenarioContext.Current.ScenarioInfo.Title.Equals("Show Plants"))
            {
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
            }
            return true;
        }

        public bool textBoxWorkCenterSearch(string code)
        {
            getCode = code;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonWorkCenter));
            if (objectDisplay.Contains(false)) { return false; }
            buttonWorkCenter.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__dialog0-searchField-I")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxSearchCodeCenter));
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(300);
            textBoxSearchCodeCenter.SendKeys(code);
            Thread.Sleep(300);
            textBoxSearchCodeCenter.SendKeys(Keys.Enter);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMList.sapMListBGSolid")));
            IWebElement listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
            var rows = listWorkCenter.FindElements(By.TagName("li"));
            foreach (var row in rows)
            {
                var stringSplit = row.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codeItemScreen = stringSplit[0];
                if (code.Equals(codeItemScreen))
                {
                    row.Click();
                }
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
            ((IJavaScriptExecutor)SetUp.Driver).ExecuteScript("arguments[0].scrollIntoView(true);", scrollDown);
            return true;
        }

        public bool buttonFillTime()
        {
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--hourFromId-inner")));
            getTimeStart = "0600";
            getTimeFinish = "'1200";
            texBoxTimeHourStart.Click();
            Thread.Sleep(1000);
            texBoxTimeHourStart.SendKeys(getTimeStart);
            Thread.Sleep(1000);
            texBoxTimeHourFinish.Click();
            Thread.Sleep(1000);
            texBoxTimeHourFinish.SendKeys(getTimeFinish);
            return true;
        }

        public bool buttonFillTimeNotOverlaping()
        {
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#__xmlview0--hourFromId-inner")));
            getTimeStart = "1300";
            getTimeFinish = "'1800";
            texBoxTimeHourStart.Click();
            Thread.Sleep(1000);
            texBoxTimeHourStart.SendKeys(getTimeStart);
            Thread.Sleep(1000);
            texBoxTimeHourFinish.Click();
            Thread.Sleep(1000);
            texBoxTimeHourFinish.SendKeys(getTimeFinish);
            return true;
        }

        public bool buttonFillTimeOverlaping()
        {
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#__xmlview0--hourFromId-inner")));
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Deleting two or more capacities"))
            {
                getTimeStart = "1900";
                getTimeFinish = "'2200";
            }
            else
            {
                getTimeStart = "0700";
                getTimeFinish = "'1400";
            }
            texBoxTimeHourStart.Click();
            Thread.Sleep(1000);
            texBoxTimeHourStart.SendKeys(getTimeStart);
            Thread.Sleep(1000);
            texBoxTimeHourFinish.Click();
            Thread.Sleep(1000);
            texBoxTimeHourFinish.SendKeys(getTimeFinish);
            return true;
        }

        public bool saveIndividualCapacity()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__button10")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSave));
            if (objectDisplay.Contains(false)) { return false; }
            buttonSave.Click();
            Thread.Sleep(1000);
            ((IJavaScriptExecutor)SetUp.Driver).ExecuteScript("arguments[0].scrollIntoView(true);", scrollDown);
            return true;
        }

        public bool selectIndividualCapacity()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--capacityTableId-table")));
            IList<IWebElement> listCapacity = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableRow.sapUiTableContentRow.sapUiTableTr"));
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Deleting a capacity":
                    listCapacity[0].Click();
                    break;
                case "Deleting two or more capacities":
                    listCapacity[0].Click();
                    listCapacity[1].Click();
                    break;
            }
            Thread.Sleep(700);
            return true;
        }

        public bool deleteIndividualCapacity()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__button9")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDelete));
            if (objectDisplay.Contains(false)) { return false; }
            buttonDelete.Click();
            Thread.Sleep(700);
            return true;
        }

        public bool deleteIndividualCapacityConfirmation()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMBtnBase.sapMBtn.sapMDialogBeginButton.sapMBarChild")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonDeleteConfirmation));
            if (objectDisplay.Contains(false)) { return false; }
            buttonDeleteConfirmation.Click();
            Thread.Sleep(700);
            return true;
        }

        public bool buttonEditIndividualCapacity()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMMenuBtn.sapMMenuBtnRegular.siggaTableButton")));
            IList<IWebElement> listCapacity = SetUp.Driver.FindElements(By.CssSelector(".sapMMenuBtn.sapMMenuBtnRegular.siggaTableButton"));
            objectDisplay.Add(validacao.ValidaElemVisivel(listCapacity[0]));
            if (objectDisplay.Contains(false)) { return false; }
            listCapacity[0].Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[contains(text(),'Edit')]")));
            buttonInternalFromTableEdit.Click();
            Thread.Sleep(700);
            return true;
        }

        public bool buttonCopyIndividualCapacity()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMMenuBtn.sapMMenuBtnRegular.siggaTableButton")));
            IList<IWebElement> listCapacity = SetUp.Driver.FindElements(By.CssSelector(".sapMMenuBtn.sapMMenuBtnRegular.siggaTableButton"));
            objectDisplay.Add(validacao.ValidaElemVisivel(listCapacity[0]));
            if (objectDisplay.Contains(false)) { return false; }
            listCapacity[0].Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[contains(text(),'Copy')]")));
            buttonInternalFromTableCopy.Click();
            Thread.Sleep(700);
            return true;
        }

        public bool editIndividualCapacity()
        {
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#editDialog--hourEditFromId-inner")));
            textBoxTimeStartEditCapacity.SendKeys("1000");
            textBoxTimeFinishEditCapacity.SendKeys("2300");
            textBoxNoteEditCapacity.SendKeys("Automação de Testes");
            buttonOkEditCapacity.Click();
            Thread.Sleep(700);
            return true;
        }

        public bool seachPersonCloneCapacity()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMSFI")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxSearchCloneCapacity));
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Copy a capacity with overlap"))
            {
                textBoxSearchCloneCapacity.SendKeys(getOnePerson);
            }
            else { textBoxSearchCloneCapacity.SendKeys(getOnePersonCopyCapacity); }
            textBoxSearchCloneCapacity.SendKeys(Keys.Enter);
            IList<IWebElement> personRow = SetUp.Driver.FindElements(By.CssSelector("#cloneDialog--personTableId-rows-row0"));
            var stringSplit = personRow[0].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string number = stringSplit[1];
            IWebElement radioBox = SetUp.Driver.FindElement(By.CssSelector("#cloneDialog--personTableId-rowsel0"));
            if (number.Equals(getOnePersonCopyCapacity) || number.Equals(getOnePerson))
            {
                radioBox.Click();
            }
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(700);
            buttonOkCopyCapacity.Click();
            return true;
        }

        public bool selectPerson()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapUiTableRow.sapUiTableContentRow.sapUiTableTr")));
            IList<IWebElement> getPersonForCopy = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableRow.sapUiTableContentRow.sapUiTableTr"));
            var split = getPersonForCopy[1].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            getOnePersonCopyCapacity = split[0];
            IList<IWebElement> rows = SetUp.Driver.FindElements(By.CssSelector(".sapUiTableRow.sapUiTableContentRow.sapUiTableTr"));
            List<string> listScenario = new List<string>();
            listScenario.Add("Search by an specific person");
            listScenario.Add("Deleting a capacity");
            listScenario.Add("Copy a capacity");
            listScenario.Add("Copy a capacity with overlap");
            listScenario.Add("Editing a capacity");
            listScenario.Add("Editing a capacity with overlap");
            listScenario.Add("Search by an specific person in a specific period");
            listScenario.Add("Registering a capacity for one person");
            listScenario.Add("Registering a capacity for one person that already has a capacity for that period, but without overlaping of time");
            listScenario.Add("Registering a capacity for one person that already has a capacity for that period with overlaping of time");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            bool scenarioContains = false;
            for (int pos = 0; pos < rows.Count; pos++)
            {
                foreach (var item in listScenario)
                {
                    if (item.Equals(scenario))
                    {
                        var stringSplit = rows[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                        string codePersonScreen = stringSplit[0];
                        getOnePerson = codePersonScreen;
                        var radioBox = string.Format("#__xmlview0--personId--valueHelpDialog-table-rowsel{0}", pos);
                        var box = SetUp.Driver.FindElements(By.CssSelector(radioBox)).FirstOrDefault();
                        box.Click();
                        scenarioContains = true;
                        break;
                    }
                }
                if (scenarioContains.Equals(true)) { break; }
                if (pos < 2)
                {
                    var stringSplit = rows[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    string codePersonScreen = stringSplit[0];
                    getMorePerson.Add("'" + codePersonScreen + "'");
                    var radioBox = string.Format("#__xmlview0--personId--valueHelpDialog-table-rowsel{0}", pos);
                    var box = SetUp.Driver.FindElements(By.CssSelector(radioBox)).FirstOrDefault();
                    box.Click();
                    if (pos == 1) { break; }
                }
            }
            IWebElement Ok = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--personId--valueHelpDialog-ok-inner")));
            Ok.Click();
            cleanTable();
            return true;
        }

        public void cleanTable()
        {
            string listPerson = personToString();
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            switch (scenario)
            {
                case "Registering a capacity for one person":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Registering a capacity for two or more people":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityMorePerson(listPerson);
                    break;
                case "Registering a capacity for one person that already has a capacity for that period, but without overlaping of time":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Registering a capacity for one person that already has a capacity for that period with overlaping of time":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Deleting a capacity":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Deleting two or more capacities":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityMorePerson(listPerson);
                    break;
                case "Editing a capacity":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Editing a capacity with overlap":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Copy a capacity":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
                case "Copy a capacity to more than one person":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityMorePerson(listPerson);
                    break;
                case "Copy a capacity with overlap":
                    new IndividualCapacity_SapConnect().cleanTableIndividualCapacityOnePerson(getOnePerson);
                    break;
            }
        }

        public bool textBoxPeriodLeft()
        {
            string feature = FeatureContext.Current.FeatureInfo.Title;
            switch (feature)
            {
                case "IndividualCapacity - Filter":
                    days = DateTime.Now.AddYears(-1);
                    calendarStart = "#__xmlview0--";
                    calendarMiddle = "periodFromId";
                    break;
                case "IndividualCapacity - Manage":
                    days = DateTime.Now;
                    if (ScenarioStepContext.Current.StepInfo.Text.Equals("I select a date from dialog capacity"))
                    {
                        calendarStart = "#cloneDialog--";
                        calendarMiddle = "periodFromId";
                    }
                    else
                    {
                        calendarStart = "#__xmlview0--";
                        calendarMiddle = "allocationFromId";
                    }
                    break;
            }
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
            IWebElement textBoxStartDate = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-icon")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxStartDate));
            Thread.Sleep(200);
            textBoxStartDate.Click();
            IWebElement buttonMonthLeft = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonMonthLeft));
            buttonMonthLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            IWebElement buttonYearLeft = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-cal--Head-B2")));
            buttonYearLeft.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearStart))
                {
                    yearSelect.Click();
                    break;
                }
            }
            IWebElement textBoxStartDateClick = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-icon")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxStartDateClick));
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxStartDateClick).DoubleClick().Perform();
            Thread.Sleep(1000);
            string firstDay = dayStart.Substring(0, 1);
            string secondDay = dayStart.Substring(1, 1);
            string firstMonth = monthStart.Substring(0, 1);
            string secondMonth = monthStart.Substring(1, 1);
            if (dayStart.StartsWith("0")) { dayStart = dayStart.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {

                var day = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--Month0-2021{0}{1}{2}{3}", firstMonth, secondMonth, firstDay, secondDay);
                var daySelect = SetUp.Driver.FindElement(By.CssSelector(day));
                if (daySelect.GetAttribute("innerText").Equals(dayStart))
                {
                    daySelect.Click();
                    break;
                }
            }
            switch (feature)
            {
                case "IndividualCapacity - Filter":
                    textBoxPeriodStartAlterPeriod.Clear();
                    string send = dayStart + "/" + monthStart + "/" + yearStart;
                    textBoxPeriodStartAlterPeriod.SendKeys(send);
                    break;
                case "IndividualCapacity - Manage":
                    if (ScenarioStepContext.Current.StepInfo.Text.Equals("I select a date from dialog capacity"))
                    {
                        textBoxPeriodStartCloneCapacity.Clear();
                        send = dayStart + "/" + monthStart + "/" + yearStart;
                        textBoxPeriodStartCloneCapacity.SendKeys(send);
                    }
                    else
                    {
                        textBoxPeriodStartAlterAlocation.Clear();
                        send = dayStart + "/" + monthStart + "/" + yearStart;
                        textBoxPeriodStartAlterAlocation.SendKeys(send);
                    }
                    break;
            }


            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool textBoxPeriodRight()
        {
            string feature = FeatureContext.Current.FeatureInfo.Title;
            switch (feature)
            {
                case "IndividualCapacity - Filter":
                    days = DateTime.Now;
                    calendarStart = "#__xmlview0--";
                    calendarMiddle = "periodToId";
                    break;
                case "IndividualCapacity - Manage":
                    days = DateTime.Now.AddDays(7);
                    if (ScenarioStepContext.Current.StepInfo.Text.Equals("I select a date to dialog capacity"))
                    {
                        calendarStart = "#cloneDialog--";
                        calendarMiddle = "periodToId";
                    }
                    else
                    {
                        calendarStart = "#__xmlview0--";
                        calendarMiddle = "allocationToId";
                    }
                    break;
            }
            end = days.ToString("dd/MM/yyyy");
            convertCalentarMonthPopPup(end);
            string dayFinish = end.Substring(0, 2);
            string yearFinish = end.Substring(6, 4);
            var dateFinish = DateTime.ParseExact(end, "dd/MM/yyyy", null);
            getFinishDate = dateFinish.ToString("yyyy/MM/dd").Replace("/", "");
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            IWebElement textBoxFinishDate = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-icon")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxFinishDate));
            Thread.Sleep(200);
            textBoxFinishDate.Click();
            IWebElement buttonMonthRight = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-cal--Head-B1")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonMonthRight));
            buttonMonthRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalMonthPicker")));
            for (int pos = 0; pos <= 11; pos++)
            {
                var month = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--MP-m{0}", pos);
                var monthSelect = SetUp.Driver.FindElement(By.CssSelector(month));
                if (monthSelect.GetAttribute("innerText").Equals(getMouth))
                {
                    monthSelect.Click();
                    break;
                }
            }
            IWebElement buttonYearRight = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-cal--Head-B2")));
            buttonYearRight.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("sapUiCalYearPicker")));
            for (int pos = 0; pos <= 19; pos++)
            {
                var year = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--YP-y202{0}0101", pos);
                var yearSelect = SetUp.Driver.FindElement(By.CssSelector(year));
                if (yearSelect.GetAttribute("innerText").Equals(yearFinish))
                {
                    yearSelect.Click();
                    break;
                }
            }
            IWebElement textBoxStartDateClick = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-icon")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxStartDateClick));
            Actions action = new Actions(SetUp.Driver);
            action.MoveToElement(textBoxStartDateClick).DoubleClick().Perform();
            Thread.Sleep(1000);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("" + calendarStart + "" + calendarMiddle + "-cal--Month0-WH6")));
            string first = dayFinish.Substring(0, 1);
            string second = dayFinish.Substring(1, 1);
            if (dayFinish.StartsWith("0")) { dayFinish = dayFinish.Remove(0, 1); }
            for (int pos = 0; pos <= 35; pos++)
            {
                var day = string.Format("" + calendarStart + "" + calendarMiddle + "-cal--Month0-202103{0}{1}", first, second);
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

        public bool validateMsgPopUpExist()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__error0")));
            objectDisplay.Add(validacao.ValidaElemVisivel(popUpExist));
            if (objectDisplay.Contains(false)) { return false; }
            Thread.Sleep(200);
            return true;
        }

        public bool validatePlantsValidate()
        {
            new IndividualCapacity_SapConnect().tablePlants();
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            IWebElement listBox = SetUp.Driver.FindElement(By.CssSelector(".sapMSelectList.sapMSltList-CTX"));
            var rows = listBox.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = IndividualCapacity_PlantExcel.XlsxInput;
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

        public bool validateWorkCentersRegistered()
        {
            new IndividualCapacity_SapConnect().tableLList(getPlant);
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMList.sapMListBGSolid")));
            IWebElement listWorkCenter = SetUp.Driver.FindElement(By.CssSelector(".sapMList.sapMListBGSolid"));
            var rows = listWorkCenter.FindElements(By.TagName("li"));
            ExcelWorksheet xlsxInput = IndividualCapacity_LListExcel.XlsxInput;
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
                string codeItemScreen = stringSplit[0];
                string descriptionScreen = stringSplit[1];
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

        public bool validatePeople()
        {
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Show people by work center"))
            {
                new IndividualCapacity_SapConnect().tableLPer_Wkc(getPlant, getCode);
            }
            else
            {
                new IndividualCapacity_SapConnect().tableLPerson(getPlant, getCode);
            }
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapUiTableRow.sapUiTableContentRow.sapUiTableTr")));
            IWebElement listPeople = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--personId--valueHelpDialog-table-table"));
            var rows = listPeople.FindElements(By.TagName("tr"));
            if (ScenarioContext.Current.ScenarioInfo.Title.Equals("Show people by work center"))
            {
                xlsxInputPeople = IndividualCapacity_LperWkcExcel.XlsxInput;
            }
            else
            {
                xlsxInputPeople = IndividualCapacity_LPersonExcel.XlsxInput;
            }
            var planilha = xlsxInputPeople;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 1; pos <= qtdRowSheet; pos++)
            {

                int posTable = pos > 5 ? 5 : pos;
                excelLPerson = planilha.Cells[linha, 1]?.Value?.ToString();
                excelLPersonUsername = planilha.Cells[linha, 2]?.Value?.ToString();
                if (string.IsNullOrEmpty(excelLPerson)) { break; }
                var stringSplit = rows[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string codePersonScreen = stringSplit[0];
                string usernamePersonScreen = stringSplit[1];
                if (excelLPerson.Equals(codePersonScreen) && excelLPersonUsername.Equals(usernamePersonScreen))
                {
                    new Util().HighlightElementPassou(rows[posTable]);
                    if (posTable < 5 || pos == qtdRowSheet) ;
                }
                else
                {
                    new Util().HighlightElementFalhou(rows[posTable]);
                    captureErrorLPerson = codePersonScreen;
                    captureErrorLPersonUsername = usernamePersonScreen;
                    return false;
                }
                linha++;
                if (posTable == 5)
                {
                    rowPeopleDown.SendKeys(Keys.Down);
                    listPeople = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--personId--valueHelpDialog-table-table"));
                    rows = listPeople.FindElements(By.TagName("tr"));
                    continue;
                }
            }
            return true;
        }

        public bool validateIndividualCapacity()
        {
            string listPerson = personToString();
            new IndividualCapacity_SapConnect().tableIndivCapa(getPlant, getCode, getOnePerson, listPerson, getStartDate, getFinishDate);
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#__xmlview0--capacityTableId-table")));
            ExcelWorksheet xlsxInput = IndividualCapacity_IndivCapExcel.XlsxInput;
            var planilha = xlsxInput;
            if (planilha is null)
            {
                IWebElement noData = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--capacityTableId-noDataMsg"));
                new Util().HighlightElementFalhou(noData);
                return false;
            }
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            List<IndividualCapacity_CellsDTO> sheetObject = new List<IndividualCapacity_CellsDTO>();
            for (int x = 2; x <= qtdRowSheet; x++)
            {
                sheetObject.Add(new IndividualCapacity_CellsDTO
                {
                    PersonalNumber = planilha.Cells[x, 1]?.Value?.ToString(),
                    Username = planilha.Cells[x, 2]?.Value?.ToString().Replace(" ", ""),
                    InitialDate = planilha.Cells[x, 3]?.Value?.ToString(),
                    EndDate = planilha.Cells[x, 4]?.Value?.ToString(),
                    StartTime = planilha.Cells[x, 5]?.Value?.ToString(),
                    Finish = planilha.Cells[x, 6]?.Value?.ToString(),
                    Capacity = planilha.Cells[x, 7]?.Value?.ToString(),
                    Note = planilha.Cells[x, 8]?.Value?.ToString().Replace(" ", ""),
                    index = x - 2
                });
            }
            int column = 9;
            for (int pos = 0; pos < qtdRowSheet - 1; pos++)
            {
                bool highlight = pos < 5 || pos == qtdRowSheet - 1;
                var sheetRows = sheetObject;
                for (int i = 1; i < column; i++)
                {
                    var columnCalendar = string.Format("#__xmlview0--capacityTableId-rows-row{0}-col{1}",
                        pos > 5 ? 5 : pos, i);
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

        private string sheetObjectValueByIndex(IndividualCapacity_CellsDTO sheetObject, int columnIndex)
        {
            switch (columnIndex)
            {
                case 1:
                    return sheetObject.PersonalNumber;
                case 2:
                    return sheetObject.Username;
                case 3:
                    return sheetObject.InitialDate;
                case 4:
                    return sheetObject.EndDate;
                case 5:
                    return sheetObject.StartTime;
                case 6:
                    return sheetObject.Finish;
                case 7:
                    return sheetObject.Capacity;
                case 8:
                    return sheetObject.Note;
                default:
                    return null;
            }
        }

        private void ColumnRowsNavigate(IWebElement nextColumn, int pos, int columnIndex)
        {
            if (columnIndex < 8)
            {
                nextColumn.SendKeys(Keys.ArrowRight);
            }
            else
            {
                IWebElement leftColumn = null;
                for (int j = 8; j >= 0; j--)
                {
                    var leftColumnCalendar = string.Format("#__xmlview0--capacityTableId-rows-row{0}-col{1}",
                        pos > 5 ? 5 : pos, j);
                    leftColumn = SetUp.Driver.FindElements(By.CssSelector(leftColumnCalendar)).FirstOrDefault();
                    leftColumn.SendKeys(Keys.ArrowLeft);
                }
                leftColumn.SendKeys(Keys.ArrowDown);
            }
        }

    }
}
