Feature: IndividualCapacity - Manage
  As an user of Sigga Planning & Scheduling
           I want to manage the individual capacity
           So that I can control the people availability 

Background:
	Given I have access to the Individual Capacity module using the login 'SIGGA127'
	When I select a plant '1000'
	And I have selected a work center 'MEC-L1'
	And I the data shown should be limited by the Data Limit '100'

Scenario: Registering a capacity for one person
	When I ask the system to show the people
	And I select a person
	And I select a date from
	And I select a date to
	#In the future and the end date should be > the start date
	And I fill in the time field
	And I click the save button
	Then The system should register the individual capacity in the table
#	#the individual capacity is registered in the /SSCN/INDIV_CAPA
#	 the new capacity should be shown on the grid as I've registered it

Scenario: Registering a capacity for two or more people
    When I ask the system to show the people 
	And I select more then one person
	And I select a date from
	And I select a date to
	#In the future and the end date should be > the start date
	And I fill in the time field
	And I click the save button
	Then The system should register the individual capacity in the table
#the individual capacity is registered in the /SSCN/INDIV_CAPA
#the new capacities should be shown on the grid as I've registered it

Scenario: Registering a capacity for one person that already has a capacity for that period, but without overlaping of time
    When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	#In the future and the end date should be > the start date
	And I fill in the time field	
	And I click the save button
	And I fill in the time without overlaping the time already registered
	And I click the save button
	Then The system should register the individual capacity in the table
#the individual capacity is registered in the /SSCN/INDIV_CAPA

Scenario: Registering a capacity for one person that already has a capacity for that period with overlaping of time
   When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	#In the future and the end date should be > the start date
	And I fill in the time field	
	And I click the save button
	And I fill in the time without overlaping the time already registered
	And I click the save button
	And Again I insert the time to override the time already registered
	And I click the save button
	Then The system should inform me that that person already has a capacity registered for that period

Scenario: Deleting a capacity
	When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I fill in the time without overlaping the time already registered
	And I click the save button
	When I select one capacity
	And I delete it
	And I confirm it
	Then The system should register the individual capacity in the table
	
Scenario: Deleting two or more capacities
    When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I fill in the time without overlaping the time already registered
	And I click the save button
	When I have already registered more than one capacity
	And I delete it
	And I confirm it
	Then The system should register the individual capacity in the table

Scenario: Editing a capacity
    When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I click on edit
	And I change the information and then click OK for confirmation
	Then The system should register the individual capacity in the table

Scenario: Editing a capacity with overlap
    When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I fill in the time without overlaping the time already registered
	And I click the save button
	And I click on edit
	And I change the information and then click OK for confirmation
	Then The system should inform me that that person already has a capacity registered for that period

Scenario: Copy a capacity
	When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I click on copy
	And I select a date from dialog capacity
	And I select a date to dialog capacity	
	And I select a person to copy and then click OK for confirmation
	Then The system should copy all the registers for the first person to the second in that period
	#the individual capacity is registered in the /SSCN/INDIV_CAPA

Scenario: Copy a capacity to more than one person
	When I ask the system to show the people 
	And I select more then one person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I click on copy
	And I select a date from dialog capacity
	And I select a date to dialog capacity	
	And I select a person to copy and then click OK for confirmation
	Then The system should copy all the registers for the first person to the second in that period

#the individual capacity is registered in the /SSCN/INDIV_CAPA
Scenario: Copy a capacity with overlap
	When I ask the system to show the people 
	And I select a person
	And I select a date from
	And I select a date to
	And I fill in the time field	
	And I click the save button
	And I click on copy
	And I select a date from dialog capacity
	And I select a date to dialog capacity	
	And I select a person to copy and then click OK for confirmation
	Then The system should inform me that that person already has a capacity registered for that period