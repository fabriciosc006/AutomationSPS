Feature: MDC - Using Filter
	 As an user of Sigga Planning & Scheduling
           I want to search for the machine downtime calendar
           So that I can know when an asset is scheduled to be stopped

Background:
	Given I have access to the Machine Downtime calendar using the login 'SIGGA127'
	 When I have selected a plant '1000'		
	And the grids should be limited by the Data Limit '1000'

#Bug
Scenario: Filtering only by plant   
	When I search it
	And I should see all the locations in that plant
	#Verify the Locations in /SSCN/LFUNC_LOC AS A LEFT JOIN /SSCN/LFUN_LOC_T with MAINTPLANT=selected
	And I should see all the equipments in that plant
	#Verify the equipments in /SSCN/LEQP with MAINTPLANT=Selected plant
	Then I should see all the machine downtime registered to those assets
	#Verify the MDC registered in /SSCN/MACH_CAL with PLANT=selected plant
	
#Bug
Scenario: Filtering by plant and work center 
	When I ask the system to show the Work Centers
	And I search for a work center code 'MECHANIK'
	And I search it
	And I should see all the locations in that plant
	#Verify the Locations in /SSCN/LFUNC_LOC with MAINTPLANT=selected, WORK_CTR=selected work center and verify the equipments in /SSCN/LEQP with MAINTPLANT=Selected plant  and WORK_CTR=selected work center
	And I should see all the equipments in that plant
	#Verify the MDC registered in /SSCN/MACH_CAL with PLANT=selected plant and WORK_CENTER= selected work center
	Then I should see all the machine downtime registered to those assets

Scenario: Filtering by plant and Functional location
	When I use the LIKE query to return the functional location code that starts with the string '00'
	And I search it
	And I should see all the locations in that plant
	Then I should see all the machine downtime registered to those assets

Scenario: Filtering by plant and Equipment	
	When I use the LIKE query to return the equipment code that starts with the string 'P-1000'
	And I search it
	And I should see all the equipments in that plant
	Then I should see all the machine downtime registered to those assets

Scenario: Filtering by plant and date	
	When I inform the initial period and the final period
	And I search it
	Then I should see all the machine downtime registered to those assets
	#In /SSCN/MACH_CAL with START_DATE between From and To

Scenario: Filtering by plant and an invalid date	
	When I inform the initial period greater than the final period
	And I search it
	And the fields with the period should be red
	Then I should see a message informing that the period is invalid
	