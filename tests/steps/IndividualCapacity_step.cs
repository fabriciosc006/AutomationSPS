using NUnit.Framework;
using SiggaPS.tests.pages;
using System.Collections.Generic;
using TechTalk.SpecFlow;

namespace SiggaPS.tests.steps
{
    [Binding]
    class IndividualCapacity_step
    {
        public string login { get; set; }
        public string password { get; set; }

        IndividualCapacity_action capacity = new IndividualCapacity_action();

        [Given(@"I have access to the Individual Capacity module using the login '(.*)'")]
        public void GivenIHaveAccessToTheIndividualCapacityModuleUsingTheLogin(string nome)
        {
            login = nome;
            password = "123690";
            if (FeatureContext.Current.FeatureInfo.Title.Equals("IndividualCapacity - Filter"))
            {
                bool _result = capacity.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=100&sap-language=EN#/individualCapacity");
                Assert.IsTrue(_result, " The page cannot be initialized  ", null);
            }
            else
            {
                bool _result = capacity.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=500&sap-language=EN#/individualCapacity");
                Assert.IsTrue(_result, " The page cannot be initialized  ", null);
            }

        }

        [Then(@"I should see all the plants registered for my user in this page")]
        public void ThenIShouldSeeAllThePlantsRegisteredForMyUserInThisPage()
        {
            bool _result = capacity.validatePlantsValidate();
            Assert.IsTrue(_result, "The SAP plant " + capacity.captureErrorPlant + " data is not shown on the screen  ", null);
        }

        [When(@"I select a plant '(.*)'")]
        public void WhenIHaveSelectedAPlant(string plant)
        {
            bool _result = capacity.selectPLant(plant);
            Assert.IsTrue(_result, "Plant request does not match in combobox  ", null);
        }

        [When(@"I click on the WorkCenter button")]
        public void WhenIClickOnTheWorkCenterButton()
        {
            bool _result = capacity.openWorkCenterButton();
            Assert.IsTrue(_result, "The button select work center is not response  ", null);
        }

        [Then(@"I must see all the work centers registered in the plant that I selected")]
        public void ThenIMustSeeAllTheWorkCentersRegisteredInThePlantThatISelected()
        {
            bool _result = capacity.validateWorkCentersRegistered();
            Assert.IsTrue(_result, "The code_item " + capacity.captureErrorLListCodeItem + " and description " + capacity.captureErrorLListCodeItemDescription + " of the LList table is different what is on the screen.  ", null);
        }

        [When(@"I ask the system to show the people")]
        public void WhenIAskTheSystemToShowThePeople()
        {
            bool _result = capacity.openPeopleButton();
            Assert.IsTrue(_result, "The button select people is not response  ", null);
        }

        [Then(@"I should see all the people registered to the plant I have selected")]
        public void ThenIShouldSeeAllThePeopleRegisteredToThePlantIHaveSelected()
        {
            bool _result = capacity.validatePeople();
            Assert.IsTrue(_result, "The person " + capacity.captureErrorLPerson + " and username " + capacity.captureErrorLPersonUsername + " of the LPerson table is different what is on the screen.  ", null);
        }

        [When(@"I have selected a work center '(.*)'")]
        public void WhenIHaveSelectedAWorkCenter(string code)
        {
            List<string> listScenario = new List<string>();
            listScenario.Add("Show Plants");
            listScenario.Add("Show Work Centers");
            listScenario.Add("ScenarioInfo.Title;");
            listScenario.Add("Search by plant");
            listScenario.Add("Search by plant in an specific period");
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
                bool _result = capacity.textBoxWorkCenterSearch(code);
                Assert.IsTrue(_result, "The name " + capacity.getCode + " has not been entered or does not exist in the database  ", null);
            }
        }

        [Then(@"I should see all the people registered to the work center I have selected")]
        public void ThenIShouldSeeAllThePeopleRegisteredToTheWorkCenterIHaveSelected()
        {
            bool _result = capacity.validatePeople();
            Assert.IsTrue(_result, "The person " + capacity.captureErrorLPerson + " and username " + capacity.captureErrorLPersonUsername + " of the LPerson table is different what is on the screen.  ", null);
        }

        [When(@"I the data shown should be limited by the Data Limit '(.*)'")]
        public void WhenITheDataShownShouldBeLimitedByTheDataLimit(string datalimit)
        {
            if (!ScenarioContext.Current.ScenarioInfo.Title.Equals("Show Plants"))
            {
                bool _result = capacity.changeDataLimit(datalimit);
                Assert.IsTrue(_result, "It was not possible to insert the data limit  ", null);
            }
        }

        [When(@"I click Search")]
        public void GivenIClickSearch()
        {
            bool _result = capacity.buttonSearch();
            Assert.IsTrue(_result, "Search button did not respond after clicking  ", null);
        }

        [Then(@"I should see all the individual capacity registered for that plant")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForThatPlant()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for people from that work center")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForPeopleFromThatWorkCenter()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [When(@"I select a person")]
        public void WhenISelectAPerson()
        {
            bool _result = capacity.selectPerson();
            Assert.IsTrue(_result, "It was not possible to select the person.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for that person")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForThatPerson()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [When(@"I select more then one person")]
        public void WhenISelectMoreThenOnePerson()
        {
            bool _result = capacity.selectPerson();
            Assert.IsTrue(_result, "It was not possible to select the person.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for each person selected")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForEachPersonSelected()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [When(@"I select a date from")]
        public void WhenISelectADateFrom()
        {
            bool _result = capacity.textBoxPeriodLeft();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [When(@"I select a date to")]
        public void WhenISelectADateTo()
        {
            bool _result = capacity.textBoxPeriodRight();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [When(@"I select a date from dialog capacity")]
        public void WhenISelectADateFromDialogCapacity()
        {
            bool _result = capacity.textBoxPeriodLeft();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [When(@"I select a date to dialog capacity")]
        public void WhenISelectADateToDialogCapacity()
        {
            bool _result = capacity.textBoxPeriodRight();
            Assert.IsTrue(_result, "No date to be searched  ", null);
        }

        [Then(@"I should see all the individual capacity registered for that plant in that specific period")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForThatPlantInThatSpecificPeriod()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for people from that work center in that specific period")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForPeopleFromThatWorkCenterInThatSpecificPeriod()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for that person in that specific period")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForThatPersonInThatSpecificPeriod()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [Then(@"I should see all the individual capacity registered for each person selected in that specific period")]
        public void ThenIShouldSeeAllTheIndividualCapacityRegisteredForEachPersonSelectedInThatSpecificPeriod()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [When(@"I fill in the time field")]
        public void WhenIFillInTheTimeField()
        {
            bool _result = capacity.buttonFillTime();
            Assert.IsTrue(_result, "Time has not been filled.  ", null);
        }

        [When(@"I fill in the time without overlaping the time already registered")]
        public void WhenIFillInTheTimeWithoutOverlapingTheTimeAlreadyRegistered()
        {
            bool _result = capacity.buttonFillTimeNotOverlaping();
            Assert.IsTrue(_result, "Time has not been filled.  ", null);
        }

        [When(@"Again I insert the time to override the time already registered")]
        public void WhenAgainIInsertTheTimeToOverrideTheTimeAlreadyRegistered()
        {
            bool _result = capacity.buttonFillTimeOverlaping();
            Assert.IsTrue(_result, "Time has not been filled.  ", null);
        }

        [When(@"I click the save button")]
        public void WhenIClickTheSaveButton()
        {
            bool _result = capacity.saveIndividualCapacity();
            Assert.IsTrue(_result, "Could not save capacity information.  ", null);
        }

        [When(@"I click on edit")]
        public void WhenIClickOnEdit()
        {
            bool _result = capacity.buttonEditIndividualCapacity();
            Assert.IsTrue(_result, "Edit button is not responding.  ", null);
        }

        [When(@"I click on copy")]
        public void ThenIClickOnCopy()
        {
            bool _result = capacity.buttonCopyIndividualCapacity();
            Assert.IsTrue(_result, "Copy button is not responding.  ", null);
        }

        [When(@"I delete it")]
        public void WhenIDeleteIt()
        {
            bool _result = capacity.deleteIndividualCapacity();
            Assert.IsTrue(_result, "Could not delete capacity information.  ", null);
        }

        [When(@"I confirm it")]
        public void WhenIConfirmIt()
        {
            bool _result = capacity.deleteIndividualCapacityConfirmation();
            Assert.IsTrue(_result, "Could not delete capacity information.  ", null);
        }

        [When(@"I select one capacity")]
        public void WhenISelectOneCapacity()
        {
            bool _result = capacity.selectIndividualCapacity();
            Assert.IsTrue(_result, "It was not possible to select the capacity.  ", null);
        }

        [When(@"I have already registered more than one capacity")]
        public void WhenIHaveAlreadyRegisteredMoreThanOneCapacity()
        {
            bool _result = capacity.selectIndividualCapacity();
            Assert.IsTrue(_result, "It was not possible to select the capacity.  ", null);
        }

        [When(@"I change the information and then click OK for confirmation")]
        public void WhenIChangeTheInformationAndThenClickOKForConfirmation()
        {
            bool _result = capacity.editIndividualCapacity();
            Assert.IsTrue(_result, "Could not save edit information.  ", null);
        }

        [Then(@"The system should register the individual capacity in the table")]
        public void ThenTheSystemShouldRegisterTheIndividualCapacityInTheTable()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [Then(@"The system should inform me that that person already has a capacity registered for that period")]
        public void ThenTheSystemShouldInformMeThatThatPersonAlreadyHasACapacityRegisteredForThatPeriod()
        {
            bool _result = capacity.validateMsgPopUpExist();
            Assert.IsTrue(_result, "The person already exists pop-up does not appear on the screen.  ", null);
        }

        [Then(@"The system should copy all the registers for the first person to the second in that period")]
        public void ThenTheSystemShouldCopyAllTheRegistersForTheFirstPersonToTheSecondInThatPeriod()
        {
            bool _result = capacity.validateIndividualCapacity();
            Assert.IsTrue(_result, "The base data is different from what is shown on the screen.  ", null);
        }

        [When(@"I select a person to copy and then click OK for confirmation")]
        public void WhenISelectAPersonToCopyAndThenClickOKForConfirmation()
        {
            bool _result = capacity.seachPersonCloneCapacity();
            Assert.IsTrue(_result, "Could not find the person.  ", null);
        }
    }
}
