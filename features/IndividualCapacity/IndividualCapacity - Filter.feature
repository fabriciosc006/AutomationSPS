Feature: IndividualCapacity - Filter
 As an user of Sigga Planning & Scheduling
           I want to search for the individual capacity
           So that I can know the people availability 

Background:
	Given I have access to the Individual Capacity module using the login 'SIGGA127'
	When I select a plant '1000'
	And I have selected a work center 'MEC-L1'
	And I the data shown should be limited by the Data Limit '100'

Scenario: Show Plants
	Then I should see all the plants registered for my user in this page
#Get the list of plants from /SSCN/PLANTS with USERID=logged user

Scenario: Show Work Centers
	When I click on the WorkCenter button
	Then I must see all the work centers registered in the plant that I selected
#The work center are on the NW table /SSCN/LLIST with NOME_LISTA = CENTRO_TRAB; CENTRO = selected plant and IDIOMA = Logged language

Scenario: Show people by plant
	When I ask the system to show the people
    Then I should see all the people registered to the plant I have selected
#The people are on the NW table /SSCN/LPERSON with CDS_KEY=selected plant

Scenario: Show people by work center
	When I ask the system to show the people
	Then I should see all the people registered to the work center I have selected
#The list of technitians are from the /SSCN/LPER_WKC with ARBPL = selected work center and WERKS = selected plant

Scenario: Search by plant
	When I click Search
	Then I should see all the individual capacity registered for that plant
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant

Scenario: Search by work center
	When I click Search
	Then I should see all the individual capacity registered for people from that work center
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant
#and PERNR = people code from that work center (get PERNR from /SSCN/LPER_WKC with ARBPL = selected work center and WERKS = selected plant)

Scenario: Search by an specific person
	When I ask the system to show the people
	And I select a person
	And I click Search
	Then I should see all the individual capacity registered for that person
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant and PERNR = selected person

Scenario: Search by a group of people
    When I ask the system to show the people
	And I select more then one person
	And I click Search
	Then I should see all the individual capacity registered for each person selected
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant and PERNR = selected people

Scenario: Search by plant in an specific period
	When I select a date from
	And I select a date to
	And I click Search
	Then I should see all the individual capacity registered for that plant in that specific period
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant and START_DATE = between the date from and the date to

Scenario: Search by work center in an specific period
	When I select a date from
	And I select a date to
	And I click Search
	Then I should see all the individual capacity registered for people from that work center in that specific period
##The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant
##PERNR = people code from that work center (get PERNR from /SSCN/LPER_WKC with ARBPL = selected work center and WERKS = selected plant)
##and START_DATE = between the date from and the date to

Scenario: Search by an specific person in a specific period
	When I ask the system to show the people
	And I select a person
	And I select a date from
	And I select a date to
	And I click Search
	Then I should see all the individual capacity registered for that person in that specific period
##The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant and PERNR = selected person
##and START_DATE = between the date from and the date to

Scenario: Search by a group of people in a specific period
	When I ask the system to show the people
	And I select more then one person
	And I select a date from
	And I select a date to
	And I click Search
	Then I should see all the individual capacity registered for each person selected in that specific period
#The individual capacities registered are on the NW tabel /SSCN/INDIV_CAPA with WERKS = selected plant and PERNR = selected people
#and START_DATE = between the date from and the date to
