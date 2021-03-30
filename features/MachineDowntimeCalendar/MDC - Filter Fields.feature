Feature: MDC - Filter Fields
   As an user of Sigga Planning & Scheduling
           I want to search for the machine downtime calendar
           So that I can know when an asset is scheduled to be stopped

Background:
	Given I have access to the Machine Downtime calendar using the login 'SIGGA127'

Scenario: Show Plants
	When I ask the system to show the plants
	Then I should see all the plants registered for my user
#Get the list of plants from /SSCN/PLANTS with USERID=logged user

#Bug in 1000 and 1200
Scenario Outline: Show Work Centers
	When I have selected a plant in array <PLANT>
	And I ask the system to show the Work Centers
	Then I should see all the work centers registered to the plant I have selected
#The work center are on the NW table /SSCN/LLIST with NOME_LISTA = CENTRO_TRAB; CENTRO = selcted plant and IDIOMA = Logged language
  	Examples:
		| PLANT |
		| 1000   |
		| 1100   |
		| 1200   |

Scenario: Search for a Work Center
	When I have selected a plant '1000'
	And I ask the system to show the Work Centers
	And I search for a work center code 'MEC-L1'
	Then I should see all the work center with this code as part of theirs

Scenario: Clean the work center field
	When I have selected a plant '1000'
	And I ask the system to show the Work Centers
	And I search for a work center code 'MEC-L1'
	When I clean the work center field
	Then the system should deselect the work center I have selected