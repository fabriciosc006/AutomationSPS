Feature: AbilityMatrix_AssetClass
	As an user of Sigga Planning & Scheduling
           I want to inform a person qualification in an asset class
           So that I can register this person specialty

Background:
	Given I have access to the Ability Matrix screen with login 'SIGGA127'
	When I select a person from Ability Matrix

Scenario: Show asset classes
	When I go to the Assets Class tab
	Then I should see all the Assets Classes in the system
#The classes comes from the table /SSCN/LLIST with NOME_LISTA = CLASSES and IDIOMA = logged language

Scenario: Set a grade for an Asset Class
	When I go to the Assets Class tab
	And I set a grade for a person in an Asset Class
	And save it
	Then I should see all the Assets Classes in the system
	#The qualification grade should be updated in the table /SSCN/AM_PERSON with PERNR=Person selected, AM_TYPE=1, AM_KEY= Selected class and AM_KN = Grade selected
