Feature: AbilityMatrix_TechObjType
	       As an user of Sigga Planning & Scheduling
           I want to inform a person qualification in a technical object type
           So that I can register this person specialty

Background:
	Given I have access to the Ability Matrix screen with login 'SIGGA127'
	When I select a person from Ability Matrix
	
Scenario: Show Technical Object Types
	When I go to the Technical Object Type tab
	Then I should see all the Technical Object Types in the system
	#The object types comes from the table /SSCN/LLIST with NOME_LISTA = TIPO_OBJ and IDIOMA = logged language

Scenario: Set a grade for an Technical Object Type
	When I go to the Technical Object Type tab
	And I set a grade for a person in an Technical Object Type
	And Save it
	Then I should see all the Technical Object Types in the system
#	#The qualification grade should be updated in the table /SSCN/AM_PERSON with PERNR=Person selected, AM_TYPE=2, AM_KEY= Selected object type and AM_KN = Grade selected
	 