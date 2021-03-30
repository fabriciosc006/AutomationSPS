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
    class AbilityMatrix_AssetClass_step
    {
        public string login { get; set; }
        public string password { get; set; }

        AbilityMatrix_AssetClass_action matrix = new AbilityMatrix_AssetClass_action();

        [Given(@"I have access to the Ability Matrix screen with login '(.*)'")]
        public void GivenIHaveAccessToTheAbilityMatrixScreenWithLogin(string nome)
        {
            login = nome;
            password = "123690";
            bool _result = matrix.acesso("http://" + nome + ":" + password + "@sigbhvmnw01.sigga.corp:8030/sap/bc/ui5_ui5/sscn/vesttro/index.html?sap-system-login-basic_auth=X&sap-client=100&sap-language=EN#/abilityMatrix");
            Assert.IsTrue(_result, " The page cannot be initialized  ", null);

        }

        [When(@"I select a person from Ability Matrix")]
        public void WhenISelectAPersonFromAbilityMatrix()
        {
            bool _result = matrix.selectPerson();
            Assert.IsTrue(_result, "The person cannot be selected  ", null);
        }


        [When(@"I go to the Assets Class tab")]
        public void WhenIGoToTheAssetsClassTab()
        {
            bool _result = matrix.selectTabAssetClass();
            Assert.IsTrue(_result, "The Asset Tab cannot be selected  ", null);
        }


        [Then(@"I should see all the Assets Classes in the system")]
        public void WhenIShouldSeeAllTheAssetsClassesInTheSystem()
        {
            bool _result = matrix.validateAssetClass();
            Assert.IsTrue(_result, "The Asset Class code " + matrix.captureErrorAssetClass + " or Description " + matrix.captureErrorDescription + " is diferent from page  ", null);
        }

        [When(@"I set a grade for a person in an Asset Class")]
        public void WhenISetAGradeForAPersonInAnAssetClass()
        {
            bool _result = matrix.selectQualification();
            Assert.IsTrue(_result, "The qualification comboBox cannot be selected  ", null);
        }

        [When(@"save it")]
        public void WhenSaveIt()
        {
            bool _result = matrix.saveButton();
            Assert.IsTrue(_result, "The save button cannot response ", null);
        }
    }
}
