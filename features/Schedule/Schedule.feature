Feature: Schedule
	  As an user of Sigga Planning and Scheduling
         I want to create a schedule
         So that I can schedule the activities of my team

Background:
	Given I have access to the programming screen using the login 'SIGGA127'
	When I open the creation dialog

Scenario: Verify work centers	
	When I give the schedule a name 'AUTOMATION TEST WITH SELENIUM'
	And I open the work center selection
	Then I should see all the work centers available for my user
#The work center are on the NW table /SSCN/LLIST with NOME_LISTA = CENTRO_TRAB; CENTRO = plants the user has access (Get the list of plants from /SSCN/PLANTS with USERID=logged user) to and IDIOMA = Logged language

Scenario: Verify Valid date
	Then The Valid to date should be according to the settings
#The valid date should be the date of the next first day of the week (FIRST_DAY_WEEK From /SSCN/PARAM) plus the Default days for Schedule validity date (SCHED_VALID_TIME From /SSCN/PARAM) -1

Scenario: Verify people assignment
        When I select the work center 'PSQA-AUT'
        Then I should see all the technitians assigned to this work center
        #The list of technitians are from the /SSCN/LPER_WKC with ARBPL = selected work center and WERKS = workcenter's plant

Scenario: Create a new schedule
	When I give the schedule a name 'AUTOMATION TEST WITH SELENIUM'
	And I select the work center 'PSQA-AUT'
	And I save it
	When the new schedule should appear on the grid and receive a Plan ID
	And the automation should validate the screen data with the NW table /SSCN/BL_PLAN
	#and in the /SSCN/BL_PLAN check the CENTRO; CRIADO_POR; DESCRICAO; DATA_INICIAL, DATA_FINAL; VALIDADE_DE; VALIDADE_ATE based on the settings
	And the automation should validate the table /SSCN/BL_PLAN data with the NW table /SSCN/BL_EORG
	#In the /SSCN/BL_EORG with the plan ID check the CENTRO_TRAB and CENTRO
	Then the automation should create a baseline of operations based on the CDS
# With the Plan ID, go to the table /SSCN/BL_OPER and compare to the CDS /SSCN/LOPER
#for the same settings of the schedule (Work center and range of dates)