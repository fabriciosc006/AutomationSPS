using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using SiggaPS.tests.dataBaseSAP.Schedule;
using SiggaPS.tests.util;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SiggaPS.tests.pages
{
    class Schedule_action
    {
        public Schedule_PlantExcel CreateExcel { get; set; }
        public string descriptionName { get; set; }
        public string workCenterName { get; set; }
        public string getPlanningIDExcel { get; set; }
        public string excelPlanningID { get; set; }
        public string excelEorgPlanningID { get; set; }
        public string excelEorgCentroTrab { get; set; }
        public string excelDescription { get; set; }
        public string excelInitialDate { get; set; }
        public string excelFinalDate { get; set; }
        public string excelValidateDe { get; set; }
        public string excelValidateAte { get; set; }
        public string extrairString { get; set; }
        public string excelBLOperWorkCenter { get; set; }
        public string excelBLOperWorkPlant { get; set; }
        public string excelBLOperEarlSchStartD { get; set; }
        public string excelLOperWorkCenter { get; set; }
        public string excelLOperWorkPlant { get; set; }
        public string excelParamFirstDayWeek { get; set; }
        public string excelParamSchedValidTime { get; set; }
        public string excelLOperEarlSchStartD { get; set; }
        public string excelCentroLList { get; set; }
        public string excelCodeItemLList { get; set; }
        public string excelLperPernr { get; set; }
        public string excelLperEndda { get; set; }
        public string captureErrorBLOperAndLoper { get; set; }
        public string captureErrorPlanScreen { get; set; }
        public string captureErrorPLanAndEorg { get; set; }
        public string captureErrorCentroLList { get; set; }
        public string captureErrorCodeItemLList { get; set; }
        public string captureErrorPernLper { get; set; }
        public string getSplitOrderId { get; set; }
        public List<string> getOrderIdList = new List<string>().Distinct().ToList();
        public string getOrderIdListToString()
        {
            return String.Join(",", getOrderIdList);
        }
        public Schedule_action()
        {
            try
            {
                PageFactory.InitElements(SetUp.Driver, this);
            }
            catch (Exception) { }
        }

        [FindsBy(How = How.XPath, Using = "//button[@id='__xmlview0--openMultiLevelMenu']")]
        [CacheLookup]
        public IWebElement menu { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__item12")]
        [CacheLookup]
        public IWebElement menuProgramation { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#__item13")]
        [CacheLookup]
        public IWebElement menuProgramationLast { get; set; }

        [FindsBy(How = How.XPath, Using = "//section[@id='selectSchedulingDialog--selectSchedulingDialog-cont']")]
        [CacheLookup]
        public IWebElement scheduleModuloScreen { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@id='selectSchedulingDialog--btnCreate-content']")]
        [CacheLookup]
        public IWebElement createButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//textarea[@id='createSchedulingDialog--textAreaDescription-inner']")]
        [CacheLookup]
        public IWebElement createNameDescription { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#createSchedulingDialog--multiInputWorkCenter-vhi")]
        [CacheLookup]
        public IWebElement buttonWorkCenterSelection { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@id='createSchedulingDialog--btnSave']")]
        [CacheLookup]
        public IWebElement buttonSaveWorkCenterSelection { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapMMessageToast.sapUiSelectable.sapContrast.sapContrastPlus.sapUiUserSelectable")]
        [CacheLookup]
        public IWebElement verifyWorkCenterExist { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='createSchedulingDialog--multiInputWorkCenter-inner']")]
        [CacheLookup]
        public IWebElement textBoxSelectWorkCenter { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@id='createSchedulingDialog--datePickerValidTo-inner']")]
        [CacheLookup]
        public IWebElement textBoxDate { get; set; }

        [FindsBy(How = How.XPath, Using = "//table[@id='createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-table']")]
        [CacheLookup]
        public IWebElement tableWorkCenter { get; set; }

        [FindsBy(How = How.Id, Using = "selectSchedulingDialog--tableOwn-rows-row0-col1")]
        [CacheLookup]
        public IWebElement statusBaseline { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".sapUiTableDataCell")]
        [CacheLookup]
        public IWebElement rowSchedule { get; set; }

        [FindsBy(How = How.Id, Using = "createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-rows-row4-col1")]
        [CacheLookup]
        public IWebElement sapUiTableCellInner { get; set; }

        [FindsBy(How = How.ClassName, Using = "siggaResourceDisplay")]
        [CacheLookup]
        public IWebElement peopleAlocation { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".ui-draggable.ui-droppable")]
        [CacheLookup]
        public IWebElement cardChangePosition { get; set; }

        public bool acesso(string url)
        {
            new Util().OpenPage(SetUp.Driver, url);
            return true;
        }
        public bool openCreationDialog()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(100));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//section[@id='selectSchedulingDialog--selectSchedulingDialog-cont']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(scheduleModuloScreen));
            if (objectDisplay.Contains(false)) { return false; }
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@id='selectSchedulingDialog--btnCreate-content']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(createButton));
            if (objectDisplay.Contains(false)) { return false; }
            createButton.Click();
            return true;
        }
        public bool createNameTextArea(string name)
        {
            descriptionName = name;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//textarea[@id='createSchedulingDialog--textAreaDescription-inner']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(createNameDescription));
            if (objectDisplay.Contains(false)) { return false; }
            createNameDescription.SendKeys(name);
            return true;
        }
        public bool buttonCenterSelection()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#createSchedulingDialog--multiInputWorkCenter-vhi")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonWorkCenterSelection));
            if (objectDisplay.Contains(false)) { return false; }
            buttonWorkCenterSelection.Click();
            return true;
        }
        public bool validationWorkCenterAvailable()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            new Schedule_SapConnect().tablePlantsInvokeTableLList();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(200));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@id='createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-table']")));
            IWebElement table = SetUp.Driver.FindElement(By.XPath("//table[@id='createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-table']"));
            var rows = table.FindElements(By.TagName("tr"));
            ExcelWorksheet xlsxInput = Schedule_LListExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            var qtdRowSheet = planilha.Cells.Worksheet.Dimension.End.Row;
            for (int pos = 1; pos <= qtdRowSheet; pos++)
            {
                int posTable = pos > 5 ? 5 : pos;
                excelCentroLList = planilha.Cells[linha, 2]?.Value?.ToString();
                excelCodeItemLList = planilha.Cells[linha, 3]?.Value?.ToString();
                if (string.IsNullOrEmpty(excelCodeItemLList)) { break; }
                var stringSplit = rows[posTable].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string centroLList = extrairString = stringSplit[0];
                string codeItemLListWithoutCaracters = extrairString = stringSplit[1];
                if (excelCentroLList.Equals(centroLList) && excelCodeItemLList.Equals(codeItemLListWithoutCaracters))
                {
                    if (posTable < 5 || pos == qtdRowSheet)
                        new Util().HighlightElementPassou(rows[posTable]);
                }
                else
                {
                    new Util().HighlightElementFalhou(rows[posTable]);
                    captureErrorCentroLList = centroLList;
                    captureErrorCodeItemLList = codeItemLListWithoutCaracters;
                    return false;
                }
                linha++;
                if (posTable == 5)
                {
                    sapUiTableCellInner.SendKeys(Keys.Down);
                    table = SetUp.Driver.FindElement(By.XPath("//table[@id='createSchedulingDialog--multiInputWorkCenter--valueHelpDialog-table-table']"));
                    rows = table.FindElements(By.TagName("tr"));
                    continue;
                }
            }
            return true;
        }
        public bool selectWorkCenter(string name)
        {
            workCenterName = name;
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@id='createSchedulingDialog--multiInputWorkCenter-inner']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(textBoxSelectWorkCenter));
            if (objectDisplay.Contains(false)) { return false; }
            textBoxSelectWorkCenter.SendKeys(name);
            Thread.Sleep(1000);
            textBoxSelectWorkCenter.SendKeys(Keys.Enter);
            return true;
        }
        public bool buttonSaveCenterSelection()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[@id='createSchedulingDialog--btnSave']")));
            objectDisplay.Add(validacao.ValidaElemVisivel(buttonSaveWorkCenterSelection));
            if (objectDisplay.Contains(false)) { return false; }
            buttonSaveWorkCenterSelection.Click();
            try
            {
                WebDriverWait waitProgramation = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(5));
                waitProgramation.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".sapMMessageToast.sapUiSelectable.sapContrast.sapContrastPlus.sapUiUserSelectable")));
                if (verifyWorkCenterExist.GetAttribute("innerText").Contains("Existe a programação"))
                {
                    return false;
                }
            }
            catch (Exception) { }
            return true;
        }
        public bool waitOpenProgramation()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.TextToBePresentInElement(statusBaseline, "Scheduled"));
            objectDisplay.Add(validacao.ValidaElemVisivel(statusBaseline));
            if (objectDisplay.Contains(false)) { return false; }
            return true;
        }

        public bool waitOpenProgramationScheduleMove(string name)
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.TextToBePresentInElement(statusBaseline, "" + name + ""));
            objectDisplay.Add(validacao.ValidaElemVisivel(statusBaseline));
            if (objectDisplay.Contains(false)) { return false; }
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector("#selectSchedulingDialog--tableOwn-rows-row0-col2")));
            new Actions(SetUp.Driver).DoubleClick(SetUp.Driver.FindElement(By.CssSelector("#selectSchedulingDialog--tableOwn-rows-row0-col2"))).Perform();
            return true;
        }

        public bool dragDropBackLogInSwinline()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(50));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#sap-ui-blocklayer-popup")));
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector(".sapMListNoData")));
            IList<IWebElement> listBacklog = SetUp.Driver.FindElements(By.CssSelector(".BacklogOperation-Card"));
            IList<IWebElement> listNoData = SetUp.Driver.FindElements(By.CssSelector(".sapMListNoData"));
            IList<IWebElement> listDragDrop = SetUp.Driver.FindElements(By.CssSelector(".ui-draggable.ui-droppable"));
            var cardBacklog = listBacklog.Where(x => x.GetAttribute("innerText").Contains(workCenterName)).FirstOrDefault();
            var cardSwinline = listNoData.Where(x => x.GetAttribute("innerText").Equals("No data")).FirstOrDefault();
            for (int pos = 0; pos < listBacklog.Count;)
            {
                listBacklog = SetUp.Driver.FindElements(By.CssSelector(".BacklogOperation-Card"));
                listNoData = SetUp.Driver.FindElements(By.CssSelector(".sapMListNoData"));
                cardBacklog = listBacklog.Where(x => x.GetAttribute("innerText").Contains(workCenterName)).FirstOrDefault();
                cardSwinline = listNoData.Where(x => x.GetAttribute("innerText").Equals("No data")).FirstOrDefault();
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#sap-ui-blocklayer-popup")));
                if (listBacklog.Count == 0) { break; }
                var stringSplitBacklog = listBacklog[pos].Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string orderIdScreen = stringSplitBacklog[0].Replace(" ", "");
                getSplitOrderId = orderIdScreen.Split('-')[0];
                Actions actions = new Actions(SetUp.Driver);
                actions.DragAndDrop(cardBacklog, cardSwinline).Build().Perform();
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#sap-ui-blocklayer-popup")));
                listDragDrop = SetUp.Driver.FindElements(By.CssSelector(".ui-draggable.ui-droppable"));
                new Schedule_SapConnect().tableBLOper(getPlanningIDExcel, getSplitOrderId);
                ExcelWorksheet xlsxInput = Schedule_BLOperExcel.XlsxInput;
                var planilha = xlsxInput;
                var linha = 2;
                string orderIdExcel = planilha.Cells[linha, 1]?.Value?.ToString();
                string activityExcel = planilha.Cells[linha, 2]?.Value?.ToString();
                string excelJoinOrderActivity = orderIdExcel + "-" + activityExcel;
                string excelControlKey = planilha.Cells[linha, 4]?.Value?.ToString();
                string excelWorkCenter = planilha.Cells[linha, 5]?.Value?.ToString();
                string excelDescription = planilha.Cells[linha, 7]?.Value?.ToString();
                string excelDate = planilha.Cells[linha, 11]?.Value?.ToString();
                string excelHour = planilha.Cells[linha, 12]?.Value?.ToString();
                string excelJoinDateHour = excelDate + " - " + excelHour;
                bool exist = false;
                listDragDrop.Where(x => x.Text.Contains(excelJoinOrderActivity))
                .ToList()
                 .ForEach(row =>
                 {
                     if ((row.Text.Contains(excelWorkCenter) && row.Text.Contains(excelDescription))
                     && ((row.Text.Contains(excelJoinDateHour) && row.Text.Contains(excelControlKey))))
                     {
                         cardChangePosition = row;
                         new Util().HighlightElementPassou(cardChangePosition);
                         exist = true;
                     }
                     else
                     {
                         cardChangePosition = row;
                         new Util().HighlightElementFalhou(cardChangePosition);
                         exist = false;
                     }
                 });
                if (exist.Equals(false))
                {
                    return false;
                }
            }
            return true;
        }

        public bool dragDropSwinlineInBackLog()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(50));
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector(".ui-draggable.ui-droppable")));
            IWebElement panelBacklog = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--spliterBacklogCardView-content-0"));
            IList<IWebElement> listDragDrop = SetUp.Driver.FindElements(By.CssSelector(".ui-draggable.ui-droppable"));
            IList<IWebElement> listBacklog = SetUp.Driver.FindElements(By.CssSelector(".BacklogOperation-Card"));
            var cardDragDrop = listDragDrop.Where(x => x.GetAttribute("innerText").Contains(workCenterName)).FirstOrDefault();
            for (int pos = 0; pos < listDragDrop.Count;)
            {
                listDragDrop = SetUp.Driver.FindElements(By.CssSelector(".ui-draggable.ui-droppable"));
                panelBacklog = SetUp.Driver.FindElement(By.CssSelector("#__xmlview0--spliterBacklogCardView-content-0"));
                cardDragDrop = listDragDrop.Where(x => x.GetAttribute("innerText").Contains(workCenterName)).FirstOrDefault();
                if (listDragDrop.Count == 0) { break; }
                Actions actions = new Actions(SetUp.Driver);
                actions.DragAndDrop(cardDragDrop, panelBacklog).Build().Perform();
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#sap-ui-blocklayer-popup")));
                listBacklog = SetUp.Driver.FindElements(By.CssSelector(".BacklogOperation-Card"));
                for (int card = 0; card < listBacklog.Count();)   
                {
                    var lastList = listBacklog.Last();
                    var stringSplitBacklog = lastList.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    string orderIdBacklog = stringSplitBacklog[0].Replace(" ", "");
                    orderIdBacklog = orderIdBacklog.Split('-')[0];
                    getOrderIdList.Add(orderIdBacklog);
                    getOrderIdList.Last();
                    getSplitOrderId = getOrderIdListToString();
                    break;
                }
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("#sap-ui-blocklayer-popup")));
                listDragDrop = SetUp.Driver.FindElements(By.CssSelector(".ui-draggable.ui-droppable"));
                new Schedule_SapConnect().tableBLOper(getPlanningIDExcel, getSplitOrderId);
                ExcelWorksheet xlsxInput = Schedule_BLOperExcel.XlsxInput;
                var planilha = xlsxInput;
                var linha = 2;
                string orderIdExcel = planilha.Cells[linha, 1]?.Value?.ToString();
                string activityExcel = planilha.Cells[linha, 2]?.Value?.ToString();
                string excelJoinOrderActivity = orderIdExcel + "-" + activityExcel;
                string excelControlKey = planilha.Cells[linha, 4]?.Value?.ToString();
                string excelWorkCenter = planilha.Cells[linha, 5]?.Value?.ToString();
                string excelDescription = planilha.Cells[linha, 7]?.Value?.ToString();
                string excelDate = planilha.Cells[linha, 11]?.Value?.ToString();
                string excelHour = planilha.Cells[linha, 12]?.Value?.ToString();
                string excelJoinDateHour = excelDate + " - " + excelHour;
                listBacklog = SetUp.Driver.FindElements(By.CssSelector(".BacklogOperation-Card"));
                bool exist = false;
                listBacklog.Where(x => x.Text.Contains(excelJoinOrderActivity))
                .ToList()
                 .ForEach(row =>
                 {
                     if ((row.Text.Contains(excelWorkCenter) && row.Text.Contains(excelDescription))
                     && ((row.Text.Contains(excelJoinDateHour) && row.Text.Contains(excelControlKey))))
                     {
                         cardChangePosition = row;
                         new Util().HighlightElementPassou(cardChangePosition);
                         exist = true;
                     }
                     else
                     {
                         cardChangePosition = row;
                         new Util().HighlightElementFalhou(cardChangePosition);
                         exist = false;
                     }
                 });
                if (exist.Equals(false))
                {
                    return false;
                }
            }
            return true;
        }


        public bool validateTableBL_Plan()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            new Schedule_SapConnect().tableBLPlan(descriptionName);
            IWebElement table = SetUp.Driver.FindElement(By.XPath("//table[@id='selectSchedulingDialog--tableOwn-table']"));
            var rows = table.FindElements(By.TagName("tr"));
            ExcelWorksheet xlsxInput = Schedule_BLPlanExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            getPlanningIDExcel = excelPlanningID = planilha.Cells[linha, 1]?.Value?.ToString();
            foreach (IWebElement row in rows.Skip(1))
            {
                rowSchedule = row;
                excelDescription = planilha.Cells[linha, 2]?.Value?.ToString();
                excelInitialDate = planilha.Cells[linha, 3]?.Value?.ToString();
                excelFinalDate = planilha.Cells[linha, 4]?.Value?.ToString();
                excelValidateDe = planilha.Cells[linha, 5]?.Value?.ToString();
                excelValidateAte = planilha.Cells[linha, 6]?.Value?.ToString();
                extrairString = row.Text.ToString();
                if (String.IsNullOrEmpty(extrairString)) { break; }
                if (extrairString.Contains(excelDescription))
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(rowSchedule));
                }
                else
                {
                    new Util().HighlightElementFalhou(rowSchedule);
                    captureErrorPlanScreen = excelDescription;
                    return false;
                }
                if (extrairString.Contains(excelInitialDate))
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(rowSchedule));
                }
                else
                {
                    new Util().HighlightElementFalhou(rowSchedule);
                    captureErrorPlanScreen = excelDescription;
                    return false;
                }
                if (extrairString.Contains(excelFinalDate))
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(rowSchedule));
                }
                else
                {
                    new Util().HighlightElementFalhou(rowSchedule);
                    captureErrorPlanScreen = excelDescription;
                    return false;
                }
                if (extrairString.Contains(excelValidateDe))
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(rowSchedule));
                }
                else
                {
                    new Util().HighlightElementFalhou(rowSchedule);
                    captureErrorPlanScreen = excelDescription;
                    return false;
                }
                if (extrairString.Contains(excelValidateAte))
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(rowSchedule));
                }
                else
                {
                    new Util().HighlightElementFalhou(rowSchedule);
                    captureErrorPlanScreen = excelDescription;
                    return false;
                }
                linha++;
            }
            return true;
        }
        public bool validateTableBL_PlanAndTableEorg()
        {
            new Schedule_SapConnect().tableBLEorg(getPlanningIDExcel);
            ExcelWorksheet xlsxInput = Schedule_BLEorgExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            excelEorgPlanningID = planilha.Cells[linha, 1]?.Value?.ToString();
            excelEorgCentroTrab = planilha.Cells[linha, 2]?.Value?.ToString();
            if (!excelEorgPlanningID.Equals(getPlanningIDExcel))
            {
                captureErrorPLanAndEorg = excelEorgPlanningID;
                return false;
            }
            if (!excelEorgCentroTrab.Equals(workCenterName))
            {
                captureErrorPLanAndEorg = excelEorgCentroTrab;
                return false;
            }
            return true;
        }
        public bool validateTableBL_OperAndTableLoper()
        {
            new Schedule_SapConnect().tableBLOper(getPlanningIDExcel, getSplitOrderId);
            new Schedule_SapConnect().tableLoper(workCenterName);
            ExcelWorksheet xlsxInputBLOper = Schedule_BLOperExcel.XlsxInput;
            ExcelWorksheet xlsxInputLOper = Schedule_LoperExcel.XlsxInput;
            var planilhaBLOper = xlsxInputBLOper;
            var planilhaLOper = xlsxInputLOper;
            var linha = 2;
            excelBLOperWorkCenter = planilhaBLOper.Cells[linha, 5]?.Value?.ToString();
            excelBLOperWorkPlant = planilhaBLOper.Cells[linha, 6]?.Value?.ToString();
            excelBLOperEarlSchStartD = planilhaBLOper.Cells[linha, 11]?.Value?.ToString();
            excelLOperWorkCenter = planilhaLOper.Cells[linha, 5]?.Value?.ToString();
            excelLOperWorkPlant = planilhaLOper.Cells[linha, 6]?.Value?.ToString();
            excelBLOperEarlSchStartD = planilhaLOper.Cells[linha, 11]?.Value?.ToString();
            if (!excelBLOperWorkCenter.Equals(excelLOperWorkCenter))
            {
                captureErrorBLOperAndLoper = excelBLOperWorkCenter;
                return false;
            }
            if (!excelBLOperWorkPlant.Equals(excelLOperWorkPlant))
            {
                captureErrorBLOperAndLoper = excelBLOperWorkPlant;
                return false;
            }
            if (!excelBLOperEarlSchStartD.Equals(excelBLOperEarlSchStartD))
            {
                captureErrorBLOperAndLoper = excelBLOperEarlSchStartD;
                return false;
            }
            return true;
        }
        public bool validateTableParamDateWithScreenDate()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            new Schedule_SapConnect().tableParam();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@id='createSchedulingDialog--multiInputWorkCenter-inner']")));
            ExcelWorksheet xlsxInput = Schedule_ParamExcel.XlsxInput;
            var planilha = xlsxInput;
            var linha = 2;
            excelParamFirstDayWeek = planilha.Cells[linha, 1]?.Value?.ToString();
            int convertFirstDayWeekInt = Convert.ToInt32(excelParamFirstDayWeek);
            excelParamSchedValidTime = planilha.Cells[linha, 2]?.Value?.ToString();
            int convertSchedValidTimeInt = Convert.ToInt32(excelParamSchedValidTime);
            var justMonday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday + 7).ToString("dd");
            string justMondayToString = justMonday.ToString();
            int convertMondayInt = Convert.ToInt32(justMondayToString);
            convertMondayInt = convertMondayInt + convertSchedValidTimeInt - convertFirstDayWeekInt;
            var dateComplete = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday + 7).ToString("" + convertMondayInt + "/MM/yyyy");
            string convertdateCompleteForValidate = dateComplete.ToString();
            if (textBoxDate.GetAttribute("value").Equals(convertdateCompleteForValidate))
            {
                objectDisplay.Add(validacao.ValidaElemVisivel(textBoxDate));
            }
            else
            {
                new Util().HighlightElementFalhou(textBoxDate);
                return false;
            }
            return true;
        }
        public bool validateTableLperWkcTechnitians()
        {
            List<object> objectDisplay = new List<object>();
            Validacao validacao = new Validacao();
            new Schedule_SapConnect().tableParam();
            new Schedule_SapConnect().tableLperWkc();
            WebDriverWait wait = new WebDriverWait(SetUp.Driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@id='createSchedulingDialog--multiInputWorkCenter-inner']")));
            ExcelWorksheet xlsxInputParam = Schedule_ParamExcel.XlsxInput;
            var planilhaParam = xlsxInputParam;
            var linhaParam = 2;
            ExcelWorksheet xlsxInputLper = Schedule_LperWkcExcel.XlsxInput;
            var planilhaLper = xlsxInputLper;
            var linhaLper = 2;
            excelParamFirstDayWeek = planilhaParam.Cells[linhaParam, 1]?.Value?.ToString();
            int convertFirstDayWeekInt = Convert.ToInt32(excelParamFirstDayWeek);
            excelParamSchedValidTime = planilhaParam.Cells[linhaParam, 2]?.Value?.ToString();
            int convertSchedValidTimeInt = Convert.ToInt32(excelParamSchedValidTime);
            var justMonday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday + 7).ToString("dd");
            string justMondayToString = justMonday.ToString();
            int convertMondayInt = Convert.ToInt32(justMondayToString);
            convertMondayInt = convertMondayInt + convertSchedValidTimeInt - convertFirstDayWeekInt;
            ((IJavaScriptExecutor)SetUp.Driver).ExecuteScript("arguments[0].scrollIntoView(true);", peopleAlocation);
            IList<IWebElement> listAlocation = SetUp.Driver.FindElements(By.ClassName("siggaResourceDisplay"));
            for (int pos = 0; pos < listAlocation.Count; pos++)
            {
                peopleAlocation = listAlocation[pos];
                excelLperPernr = planilhaLper.Cells[linhaLper, 1]?.Value?.ToString();
                excelLperEndda = planilhaLper.Cells[linhaLper, 2]?.Value?.ToString();
                string cutYear = excelLperEndda.Insert(4, "/");
                string cutDay = cutYear.Insert(7, "/");
                string formatingDate = cutDay;
                formatingDate = formatingDate.Split('/')[2];
                int justDayInt = Convert.ToInt32(formatingDate);
                string indexOne = peopleAlocation.GetAttribute("innerText");
                indexOne = indexOne.Split('-')[0];
                if (indexOne.Equals(excelLperPernr) && convertMondayInt <= justDayInt)
                {
                    objectDisplay.Add(validacao.ValidaElemVisivel(peopleAlocation));
                }
                else
                {
                    new Util().HighlightElementFalhou(peopleAlocation);
                    captureErrorPernLper = excelLperPernr;
                    return false;
                }
                linhaLper++;
            }
            return true;
        }
    }
}


