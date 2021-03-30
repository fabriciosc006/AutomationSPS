using NUnit.Framework;
using SAP.Middleware.Connector;
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
    class Schedule_step
    {
        public string login { get; set; }
        public string password { get; set; }

        Schedule_action schedule = new Schedule_action();

        [Given(@"I have access to the programming screen using the login '(.*)'")]
        public void GivenIHaveAccessToTheProgrammingScreenUsingTheLogin(string nome)
        {
            RfcConfigParameters parms = new RfcConfigParameters();
            parms.Add(RfcConfigParameters.Name, "QW1");
            parms.Add(RfcConfigParameters.AppServerHost, "10.10.10.177");
            parms.Add(RfcConfigParameters.SystemNumber, "30");
            parms.Add(RfcConfigParameters.User, "SIGGA127");
            parms.Add(RfcConfigParameters.Password, "123690");
            parms.Add(RfcConfigParameters.Client, "100");
            RfcDestination rfcDest = RfcDestinationManager.GetDestination(parms);
            RfcRepository rfcRep = rfcDest.Repository;
            IRfcFunction IReader = rfcRep.CreateFunction("ZSSCN_DELETE_BL_PLAN_FOR_USER");
            IReader.Invoke(rfcDest);

            login = nome;
            password = "123690";
            bool _result = schedule.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=100&sap-language=EN#/scheduling");
            Assert.IsTrue(_result, " The page cannot be initialized  ", null);
        }

        [When(@"I open the creation dialog")]
        public void WhenIOpenTheCreationDialog()
        {
            bool _result = schedule.openCreationDialog();
            Assert.IsTrue(_result, "The button create is not response  ", null);
        }

        [When(@"I give the schedule a name '(.*)'")]
        public void WhenIGiveTheScheduleAName(string name)
        {
            bool _result = schedule.createNameTextArea(name);
            Assert.IsTrue(_result, "Does not display the name text field to insert  ", null);
        }

        [When(@"I open the work center selection")]
        public void WhenIOpenTheWorkCenterSelection()
        {
            bool _result = schedule.buttonCenterSelection();
            Assert.IsTrue(_result, "The button select work center is not response  ", null);
        }

        [Then(@"I should see all the work centers available for my user")]
        public void ThenIShouldSeeAllTheWorkCentersAvailableForMyUser()
        {
            bool _result = schedule.validationWorkCenterAvailable();
            Assert.IsTrue(_result, "The data Center "+schedule.captureErrorCentroLList+" and data Work Center " + schedule.captureErrorCodeItemLList + "  in the LList table is different from what is shown on the screen.  ", null);
        }

        [When(@"I select the work center '(.*)'")]
        public void WhenISelectTheWorkCenter(string center)
        {
            bool _result = schedule.selectWorkCenter(center);
            Assert.IsTrue(_result, "The textbox select work center is not response  ", null);
        }

        [When(@"I save it")]
        public void WhenISaveIt()
        {
            bool _result = schedule.buttonSaveCenterSelection();
            Assert.IsTrue(_result, "The baseline cannot be created, as there is programming with the same name and the same login  ", null);
        }

        [When(@"the new schedule should appear on the grid and receive a Plan ID")]
        public void ThenTheNewScheduleShouldAppearOnTheGridAndReceiveAPlanID()
        {
            bool _result = schedule.waitOpenProgramation();
            Assert.IsTrue(_result, "The baseline did not generate the schedule  ", null);
        }

        [When(@"this schedule is still in '(.*)' status")]
        public void WhenThisScheduleIsStillInStatus(string name)
        {
            bool _result = schedule.waitOpenProgramationScheduleMove(name);
            Assert.IsTrue(_result, "The baseline did not generate the schedule  ", null);
        }

        [When(@"I drag an operation from backlog and drop it in a swinline inside the scheduling window")]
        public void WhenIDragAnOperationFromBacklogAndDropItInASwinlineInsideTheSchedulingWindow()
        {
            bool _result = schedule.dragDropBackLogInSwinline();
            Assert.IsTrue(_result, "The swinline card presents different data from the base  ", null);
        }

        [Then(@"I drag back swinline cards and drop them into a backlog inside the window")]
        public void ThenIDragBackSwinlineCardsAndDropThemIntoABacklogInsideTheWindow()
        {
            bool _result = schedule.dragDropSwinlineInBackLog();
            Assert.IsTrue(_result, "The swinline card presents different data from the base  ", null);
        }

        [When(@"the automation should validate the screen data with the NW table /SSCN/BL_PLAN")]
        public void ThenTheAutomationShouldValidadeTheScreenDataWithTheNWTableSSCNBL_PLAN()
        {
            bool _result = schedule.validateTableBL_Plan();
            Assert.IsTrue(_result, "The data  " + schedule.captureErrorPlanScreen + "  data in the BL_PLAN table is different from the programming screen  ", null);
        }
        [When(@"the automation should validate the table /SSCN/BL_PLAN data with the NW table /SSCN/BL_EORG")]
        public void ThenTheAutomationShouldValidadeTheTableSSCNBL_PLANDataWithTheNWTableSSCNBL_EORG()
        {
            bool _result = schedule.validateTableBL_PlanAndTableEorg();
            Assert.IsTrue(_result, "The information in the BL_Plan and BL_Eorg tables does not match the information  ", null);
        }

        [Then(@"the automation should create a baseline of operations based on the CDS")]
        public void ThenTheAutomationShouldCreateABaselineOfOperationsBasedOnTheCDS()
        {
            bool _result = schedule.validateTableBL_OperAndTableLoper();
            Assert.IsTrue(_result, "The  " + schedule.captureErrorBLOperAndLoper + "  data in the BL_OPER table is different from the BL_LOPER table  ", null);
        }
        [Then(@"The Valid to date should be according to the settings")]
        public void ThenTheValidToDateShouldBeAccordingToTheSettings()
        {
            bool _result = schedule.validateTableParamDateWithScreenDate();
            Assert.IsTrue(_result, "The date is different from the screen for the database  ", null);
        }
        [Then(@"I should see all the technitians assigned to this work center")]
        public void ThenIShouldSeeAllTheTechnitiansAssignedToThisWorkCenter()
        {
            bool _result = schedule.validateTableLperWkcTechnitians();
            Assert.IsTrue(_result, "The  "+schedule.captureErrorPernLper+"  data of the technical allocation has a different value than the base with the screen  ", null);
        }

    }
}
