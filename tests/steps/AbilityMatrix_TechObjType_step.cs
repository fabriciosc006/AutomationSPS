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
    class AbilityMatrix_TechObjType_step
    {
        AbilityMatrix_TechObjType_action objtype = new AbilityMatrix_TechObjType_action();

        [When(@"I go to the Technical Object Type tab")]
        public void WhenIGoToTheTechnicalObjectTypeTab()
        {
            bool _result = objtype.selectTabTechObjType();
            Assert.IsTrue(_result, "The Asset Tab cannot be selected  ", null);
        }

        [Then(@"I should see all the Technical Object Types in the system")]
        public void WhenIShouldSeeAllTheTechnicalObjectTypesInTheSystem()
        {
            bool _result = objtype.validateTechObjTypeClass();
            Assert.IsTrue(_result, "The Tech Object Type code " + objtype.captureErrorAssetClass + " or Description " + objtype.captureErrorDescription + " is diferent from page  ", null);
        }

        [When(@"I set a grade for a person in an Technical Object Type")]
        public void WhenISetAGradeForAPersonInAnTechnicalObjectType()
        {
            bool _result = objtype.selectQualification();
            Assert.IsTrue(_result, "The qualification comboBox cannot be selected  ", null);
        }

        [When(@"Save it")]
        public void WhenSaveIt()
        {
            bool _result = objtype.saveButton();
            Assert.IsTrue(_result, "The save button cannot response ", null);
        }

        //[Then(@"this person should receive this grade in that Technical Object Type")]
        //public void ThenThisPersonShouldReceiveThisGradeInThatTechnicalObjectType()
        //{
        //    ScenarioContext.Current.Pending();
        //}





    }
}
