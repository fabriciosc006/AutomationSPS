Feature: MDC - Manage MDC
 As an user of Sigga Planning & Scheduling
           I want register a machine downtime
           So that the system can know when an asset is scheduled to be stopped

Background:
	Given I have access to the Machine Downtime calendar using the login 'SIGGA127'
	When I have selected a plant '1000'
	And I use the LIKE query to return the functional location code that starts with the string '00'
	And I search it

Scenario: Registering a downtime for one functional location
	When I select a functional location
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	#In the future and the end date should be > the start date
	And I fill in the time
	And I select a Downtime type
	And I save all downtime information
	Then the new downtime should be shown on the grid as I've registered it

Scenario: Registering a downtime for two (or more) functional locations
	When I select two functional locations
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	#In the future and the end date should be > the start date
	And I fill in the time
	And I select a Downtime type
	And I save all downtime information
	Then the new downtime should be shown on the grid as I've registered it

Scenario: Registering a downtime for one equipment
	When I go to the equipment tab
	And I select an equipment
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	#In the future and the end date should be > the start date
	And I fill in the time
	And I select a Downtime type
	And I save all downtime information
	Then the new downtime should be shown on the grid as I've registered it

Scenario: Registering a downtime for two (or more) equipments
	When I go to the equipment tab
	And I select two equipments
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	#In the future and the end date should be > the start date
	And I fill in the time
	And I select a Downtime type
	And I save all downtime information
	Then the new downtime should be shown on the grid as I've registered it

Scenario: Deleting one downtime
	When I select one data from the machine's calendar
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	And I fill in the time
	And I select a Downtime type
	And I delete the selected information from the calendar machine grid
	Then downtime should be shown in the grid as I deleted it

Scenario: Deleting two (or more) downtimes
	When I select two data from the machine's calendar
	And I go to the Downtime Information tab
	And I fill in the dates with a valid date
	And I fill in the time
	And I select a Downtime type
	And I delete the selected information from the calendar machine grid
	Then downtime should be shown in the grid as I deleted it

Scenario: Editing one downtime
	When I select one data from the machine's calendar
	And   I click the Edit button
	And I edit Time From as an example
	Then downtime should be shown in the grid as I edited it