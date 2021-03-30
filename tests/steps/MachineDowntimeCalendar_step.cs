using NUnit.Framework;
using SiggaPS.tests.pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.steps
{
    [Binding]
    class MachineDowntimeCalendar_step
    {
        public string login { get; set; }
        public string password { get; set; }
        MachineDowntimeCalendar_action machine = new MachineDowntimeCalendar_action();

        [Given(@"I have access to the Machine Downtime calendar using the login '(.*)'")]
        public void GivenIHaveAccessToTheMachineDowntimeCalendarUsingTheLogin(string nome)
        {
            login = nome;
            password = "123690";
            if (FeatureContext.Current.FeatureInfo.Title.Equals("MDC - Manage MDC"))
            {
                bool _result = machine.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=500&sap-language=EN#/stoppedMachineCalendar");
                Assert.IsTrue(_result, " The page cannot be initialized  ", null);
            }
            else
            {
                bool _result = machine.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=100&sap-language=EN#/stoppedMachineCalendar");
                Assert.IsTrue(_result, " The page cannot be initialized  ", null);
            }

        }

        [When(@"I ask the system to show the plants")]
        public void WhenIAskTheSystemToShowThePlants()
        {
            bool _result = machine.openPlantComboBox();
            Assert.IsTrue(_result, "The combobox plant did not respond to the command to expand  ", null);
        }

        [Then(@"I should see all the plants registered for my user")]
        public void ThenIShouldSeeAllThePlantsRegisteredForMyUser()
        {
            bool _result = machine.validatePlantsValidate();
            Assert.IsTrue(_result, "The SAP plant " + machine.captureErrorPlant + " data is not shown on the screen  ", null);
        }

        [When(@"I have selected a plant in array (.*)")]
        public void GivenIHaveSelectedAPlantInArry(string plant)
        {
            bool _result = machine.selectPLant(plant);
            Assert.IsTrue(_result, "Plant request does not match in combobox  ", null);
        }

        [When(@"I have selected a plant '(.*)'")]
        public void GivenIHaveSelectedAPlant(string plant)
        {
            bool _result = machine.selectPLant(plant);
            Assert.IsTrue(_result, "Plant request does not match in combobox  ", null);
        }

        [When(@"I ask the system to show the Work Centers")]
        public void WhenIAskTheSystemToShowTheWorkCenters()
        {
            bool _result = machine.openWorkCenterButton();
            Assert.IsTrue(_result, "Calling the Work Center button does not open  ", null);
        }

        [Then(@"I should see all the work centers registered to the plant I have selected")]
        public void ThenIShouldSeeAllTheWorkCentersRegisteredToThePlantIHaveSelected()
        {
            bool _result = machine.validateWorkCentersRegistered();
            Assert.IsTrue(_result, "The code_item " + machine.captureErrorLListCodeItem + " and description " + machine.captureErrorLListCodeItemDescription + " of the LList table is different what is on the screen.  ", null);
        }

        [When(@"I search for a work center code '(.*)'")]
        public void WhenISearchForAWorkCenterCode(string code)
        {
            bool _result = machine.textBoxWorkCenterSearch(code);
            Assert.IsTrue(_result, "The name " + machine.getCode + " has not been entered or does not exist in the database  ", null);
        }

        [When(@"I use the LIKE query to return the functional location code that starts with the string '(.*)'")]
        public void WhenIInformAFunctionalLocationCode(string code)
        {
            List<string> listScenario = new List<string>();
            listScenario.Add("Deleting one downtime");
            listScenario.Add("Deleting two (or more) downtimes");
            listScenario.Add("Editing one downtime");
            listScenario.Add("Registering a downtime for two (or more) equipments");
            listScenario.Add("Registering a downtime for one equipment");
            string scenario = ScenarioContext.Current.ScenarioInfo.Title;
            bool exist = false;
            foreach (var item in listScenario)
            {
                if (item.Equals(scenario))
                {
                    exist = true;
                    break;
                }               
            }
            if (exist.Equals(false))
            {
                bool _result = machine.textBoxFunctionalLocationSearch(code);
                Assert.IsTrue(_result, "No functional location code to be searched  ", null);
            }
        }

        [When(@"I use the LIKE query to return the equipment code that starts with the string '(.*)'")]
        public void WhenIUseTheLIKEQueryToReturnTheEquipmentCodeThatStartsWithTheString(string code)
        {
            bool _result = machine.textBoxEquipmentSearch(code);
            Assert.IsTrue(_result, "No equipment code to be searched  ", null);
        }

        [When(@"I inform the initial period and the final period")]
        public void GivenIInformTheInitialPeriodAndTheFinalPeriod()
        {
            bool _result = machine.textBoxStartDateFinishDate();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [When(@"I inform the initial period greater than the final period")]
        public void WhenIInformTheInitialPeriodGreaterThanTheFinalPeriod()
        {
            bool _result = machine.textBoxStartDateFinishDate();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [Then(@"I should see all the work center with this code as part of theirs")]
        public void ThenIShouldSeeAllTheWorkCenterWithThisCodeAsPartOfTheirs()
        {
            bool _result = machine.validateWorkCentersRegisteredScenario3();
            Assert.IsTrue(_result, "The code_item " + machine.captureErrorLListCodeItem + " and description " + machine.captureErrorLListCodeItemDescription + " of the LList table is different what is on the screen.  ", null);
        }

        [When(@"I clean the work center field")]
        public void WhenICleanTheWorkCenterField()
        {
            bool _result = machine.deleteWorkCenter();
            Assert.IsTrue(_result, "The delete button did not respond  ", null);
        }

        [Then(@"the system should deselect the work center I have selected")]
        public void ThenTheSystemShouldDeselectTheWorkCenterIHaveSelected()
        {
            bool _result = machine.validateButtonClearAndTextBox();
            Assert.IsTrue(_result, "The item has not been deleted on the screen  ", null);
        }

        [When(@"the grids should be limited by the Data Limit '(.*)'")]
        public void GivenTheGridsShouldBeLimitedByTheDataLimit(string datalimit)
        {
            bool _result = machine.changeDataLimit(datalimit);
            Assert.IsTrue(_result, "It was not possible to insert the data limit  ", null);
        }

        [When(@"I search it")]
        public void GivenISearchIt()
        {
            bool _result = machine.buttonSearch();
            Assert.IsTrue(_result, "Search button did not respond after clicking  ", null);
        }

        [When(@"I should see all the locations in that plant")]
        public void GivenIShouldSeeAllTheLocationsInThatPlant()
        {
            bool _result = machine.validateLocations();
            Assert.IsTrue(_result, "the data locations " + machine.captureErrorFuncLocLocation + " in the /SSCN/LFUNC_LOC table does is different what is on the screen.  ", null);
        }

        [When(@"I should see all the equipments in that plant")]
        public void GivenIShouldSeeAllTheEquipmentsInThatPlant()
        {
            bool _result = machine.validateEquipments();
            Assert.IsTrue(_result, "the data equipaments " + machine.captureErrorLeqpEqunr + " in the /SSCN/LEQP table is different what is on the screen.  ", null);
        }

        [When(@"the fields with the period should be red")]
        public void WhenTheFieldsWithThePeriodShouldBeRed()
        {
            bool _result = machine.fieldsPeriodWithRed();
            Assert.IsTrue(_result, "The fields are not bordered by red for error indication ", null);
        }

        [Then(@"I should see a message informing that the period is invalid")]
        public void ThenIShouldSeeAMessageInformingThatThePeriodIsInvalid()
        {
            bool _result = machine.toastErrorMessage();
            Assert.IsTrue(_result, "No error message appears on the screen ", null);
        }

        [Then(@"I should see all the machine downtime registered to those assets")]
        public void ThenIShouldSeeAllTheMachineDowntimeRegisteredToThoseAssets()
        {
            bool _result = machine.validateMachineCalendar();
            Assert.IsTrue(_result, "the data machine calendar presents different data between the base and the screen  ", null);
        }

        [Then(@"the new downtime should be shown on the grid as I've registered it")]
        public void ThenTheNewDowntimeShouldBeShownOnTheGridAsIVeRegisteredIt()
        {
            bool _result = machine.validateMachineCalendar();
            Assert.IsTrue(_result, "the data machine calendar presents different data between the base and the screen  ", null);
        }

        [Then(@"downtime should be shown in the grid as I deleted it")]
        public void ThenDowntimeShouldBeShownInTheGridAsIDeletedIt()
        {
            bool _result = machine.validateMachineCalendar();
            Assert.IsTrue(_result, "the data machine calendar presents different data between the base and the screen  ", null);
        }

        [Then(@"downtime should be shown in the grid as I edited it")]
        public void ThenDowntimeShouldBeShownInTheGridAsIEditedIt()
        {
            bool _result = machine.validateMachineCalendar();
            Assert.IsTrue(_result, "the data machine calendar presents different data between the base and the screen  ", null);
        }

        [When(@"I click the Edit button")]
        public void WhenIClickTheEditButton()
        {
            bool _result = machine.buttonEditMachine();
            Assert.IsTrue(_result, "The edit button does not respond to execution ", null);
        }

        [When(@"I edit Time From as an example")]
        public void WhenIEditTimeFromAsAnExample()
        {
            bool _result = machine.editTimeFrom();
            Assert.IsTrue(_result, "Could not enter time in Time from ", null);
        }

        [When(@"I select a functional location")]
        public void WhenISelectAFunctionalLocation()
        {
            bool _result = machine.selectFunctional();
            Assert.IsTrue(_result, "Could not choose location ", null);
        }

        [When(@"I select two functional locations")]
        public void WhenISelectTwoFunctionalLocations()
        {
            bool _result = machine.selectFunctional();
            Assert.IsTrue(_result, "Could not choose location ", null);
        }

        [When(@"I go to the equipment tab")]
        public void WhenIGoToTheEquipmentTab()
        {
            bool _result = machine.selectTabEquipments();
            Assert.IsTrue(_result, "Could not to click in the equipments tab ", null);
        }

        [When(@"I select an equipment")]
        public void WhenISelectAnEquipment()
        {
            bool _result = machine.selectEquipments();
            Assert.IsTrue(_result, "Could not choose equipments ", null);
        }

        [When(@"I select two equipments")]
        public void WhenISelectTwoEquipments()
        {
            bool _result = machine.selectEquipments();
            Assert.IsTrue(_result, "Could not choose equipments ", null);
        }

        [When(@"I select one data from the machine's calendar")]
        public void WhenISelectOneDataFromTheMachineSCalendar()
        {
            bool _result = machine.selectDataMachineCalendar();
            Assert.IsTrue(_result, "Could not choose data from the machine's calendar ", null);
        }

        [When(@"I select two data from the machine's calendar")]
        public void WhenISelectTwoDataFromTheMachineSCalendar()
        {
            bool _result = machine.selectDataMachineCalendar();
            Assert.IsTrue(_result, "Could not choose data from the machine's calendar ", null);
        }

        [When(@"I go to the Downtime Information tab")]
        public void WhenIGoToTheDowntimeInformationTab()
        {
            bool _result = machine.clickTabDowntimeInformation();
            Assert.IsTrue(_result, "The downtime information tab is not accessible ", null);
        }

        [When(@"I fill in the dates with a valid date")]
        public void WhenIFillInTheDatesWithAValidDate()
        {
            bool _result = machine.textBoxDowntimeFututeStartDateFinishDate();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [When(@"I fill in the time")]
        public void WhenIFillInTheTime()
        {
            bool _result = machine.buttonFillTime();
            Assert.IsTrue(_result, "Time has not been filled.  ", null);
        }

        [When(@"I select a Downtime type")]
        public void WhenISelectADowntimeType()
        {
            bool _result = machine.selectDowntimeType();
            Assert.IsTrue(_result, "It was not possible to select the downtime type.  ", null);
        }

        [When(@"I save all downtime information")]
        public void WhenISaveSavedAllDowntimeInformation()
        {
            bool _result = machine.saveDowntimeInformation();
            Assert.IsTrue(_result, "Could not save downtime information  ", null);
        }

        [When(@"I delete the selected information from the calendar machine grid")]
        public void WhenIDeleteTheSelectedInformationFromTheCalendarMachineGrid()
        {
            bool _result = machine.deleteDowntimeInformation();
            Assert.IsTrue(_result, "Could not delete downtime information  ", null);
        }

    }
}
