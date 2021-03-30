Feature: Schedule - Move_Operations
	  As an user of Sigga Planning and Scheduling
         I want to move my operations
         So that I can schedule the activities of my team

Background:
	Given I have access to the programming screen using the login 'SIGGA127'
    When I open the creation dialog
    And I give the schedule a name 'AUTOMATION TEST WITH SELENIUM'
	And I select the work center 'PSQA-AUT'
	And I save it
	And this schedule is still in 'Scheduled' status

Scenario: Move operation card backlog to swinline and swinline for backlog (drag and drop)
	When I drag an operation from backlog and drop it in a swinline inside the scheduling window
	Then I drag back swinline cards and drop them into a backlog inside the window
#Scenario: Move operation from a swinline to the backlog successfully(drag and drop)
#	And I have already moved an operation from the backlog to a swinline
#	When I drag an operation, that was originaly in the backlog, from a swinline
#	And drop it in the backlog
#	Then the system should move the operation to the backlog
#	And the system should change the basic start date of that operation to its original date
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the operation duration
#	And the general availability of the work center in the resources bar should be updated with the operation duration
#
#Scenario: Move operation from a swinline to the backlog invalid(drag and drop)
#	When I drag an operation, that wasn't originaly in the backlog, from a swinline
#	And drop it in the backlog
#	Then the system should show me an error
#	And the system shouldn't change anything
#
#Scenario: Move operation from backlog to a swinline (quick move)
#	When I select an operation in the backlog
#	And I open the quick move dialog
#	And I select a date in the future
#	Then the system should move the operation to that date swinline
#	And the system should change the basic start date of that operation in the baseline
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the operation duration
#	And the general availability of the work center in the resources bar should be updated with the operation duration
#
#Scenario: Move operation from a swinline to another (quick move)
#	When I select an operation in the schedule window
#	And I open the quick move dialog
#	And I select a date in the future
#	Then the system should move the operation to that date swinline
#	And the system should change the basic start date of that operation in the baseline
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the operation duration
#
#Scenario: Move more than one operations from backlog to a swinline (quick move)
#	When I select more than one operations in the backlog
#	And I open the quick move dialog
#	And I select a date in the future
#	Then the system should move the operations to that date swinline
#	And the system should change the basic start date of the operations in the baseline
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the operations duration
#	And the general availability of the work center in the resources bar should be updated with the operations duration
#
#Scenario: Move more than one operations from a swinline to another (quick move)
#	When I select more than one operation in the schedule window
#	And I open the quick move dialog
#	And I select a date in the future
#	Then the system should move the operations to that date swinline
#	And the system should change the basic start date of the operations in the baseline
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the operations duration
#
#Scenario: Move an order from one day to another in gantt chart
#	When I open the gantt chart view
#	And I drag an order bar
#	And I drop it in another day
#	Then the system should move the order and its operations to that date
#	And the system should change the basic start date of the order's operations in the baseline
#	#check table /SSCN/BL_OPER
#	And the day availability of the work center in the resources bar should be updated with the order's operations duration
#
#Scenario: Move an order's start time in gantt chart
#	When I open the gantt chart view
#	And I drag an order bar
#	And I drop it in the same day but in a different time
#	Then the system should move the order and its operations to that time
#	And the system should change the basic start time of the order's operations in the baseline
#
##check table /SSCN/BL_OPER field EARL_SCH_START_T
#Scenario: Move an operation from one day to another in gantt chart
#	When I open the gantt chart view
#	And I expand the order to see its operations
#	And I drag an operation bar
#	And I drop it in another day
#	Then the system should move the operation to that date
#	And the system should change the basic start date of the operation in the baseline
#	#check table /SSCN/BL_OPER
#	And the order bar should follow the operation bar
#	And the day availability of the work center in the resources bar should be updated with the operation duration
#
#Scenario: Move an operation's start time in gantt chart
#	When I open the gantt chart view
#	And I expand the order to see its operations
#	And I drag an operation bar
#	And I drop it in the same day but in a different time
#	Then the system should move the operation to that time
#	And the system should change the basic start time of the operation in the baseline
#	#check table /SSCN/BL_OPER
#	And the order bar should follow the operation bar