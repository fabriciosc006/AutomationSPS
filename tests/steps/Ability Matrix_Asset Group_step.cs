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
    class Ability_Matrix_Asset_Group_step
    {

        Ability_Matrix_Asset_Group_action group = new Ability_Matrix_Asset_Group_action();
        [When(@"I go to the Asset Group tab")]
        public void WhenIGoToTheAssetGroupTab()
        {
            bool _result = group.selectTabAssetGroup();
            Assert.IsTrue(_result, "The Asset Group Tab cannot be selected  ", null);
        }

        [Then(@"I should see all the asset groups registered for this user")]
        public void ThenIShouldSeeAllTheAssetGroupsRegisteredForThisUser()
        {
            Console.WriteLine("step2");
            //ScenarioContext.Current.Pending();
        }

        [When(@"I go to Equipment List")]
        public void WhenIGoToEquipmentList()
        {
            bool _result = group.gotogrouplist();
            Assert.IsTrue(_result, "The Equipment List cannot be displayed  ", null);
        }

        [When(@"I search for an equipment")]
        public void WhenISearchForAnEquipment()
        {
            bool _result = group.textBoxAssetGroupSearch("100029");
            Assert.IsTrue(_result, "The name " + group.getCode + " has not been entered or does not exist in the database  ", null);
        }

        [Then(@"I should see all the equipments with this code or part of description")]
        public void ThenIShouldSeeAllTheEquipmentsWithThisCodeOrPartOfDescription()
        {
           // ScenarioContext.Current.Pending();
        }


    }
}
