Feature: Ability Matrix - Asset Group
           As an user of Sigga Planning & Scheduling
           I want to inform a person qualification in an asset
           So that I can register this person specialty

    Background:
        Given I have access to the Ability Matrix screen with login 'SIGGA127'
        When I select a person from Ability Matrix
        
    Scenario: Show Asset groups
        When I go to the Asset Group tab
        Then I should see all the asset groups registered for this user
      #The qualification grades comes from the table /SSCN/AM_PERSON with PERNR=Person selected and AM_TYPE=3

    Scenario: Search an equipment 
        When I go to the Asset Group tab
        And I go to Equipment List 
        And I search for an equipment
    #    Then I should see all the equipments with this code or part of description
    #    #The Equipment code comes from /SSCN/LEQP_T
    #    #The Equipment description comes from /SSCN/LEQP_T with SPRAS = logged language

#    Scenario: Add an equipment
#        When I go to the Asset Group tab
#        And I go to Equipment List 
#        And I search for an equipment
#        And I add it
#        And save it
#        Then I should see this equipment in the grid with less grade
#        #AM_KN = 1 AM_KEY_TYPE = E
#
#    Scenario: Search a functional location
#        When I go to the Asset Group tab
#        And I go to the functional location list
#        And I search for a functional location
#        Then I should see all the functional locations with this term as a code
#        And I should see all the functional locations with this term as part of their code
#        #The Functional Location code comes from /SSCN/LFUNC_LOC_T
#        And I should see all the functional locations with this term as a description
#        And I should see all the functional locations with this term as part of their description
#        #The Functional Location description comes from /SSCN/LFUN_LOC_T with SPRAS = logged language
#
#    Scenario: Add a functional location
#        When I go to the Asset Group tab
#        And I go to the functional location list
#        And I search for a functional location
#        And I add it
#        Then I should see this functional location in the grid with the lowest enabler grade possible
#        #AM_KN = 1 AM_KEY_TYPE = L
#
#    Scenario: Set a grade for an asset 
#        When I go to the Asset Group tab
#        And I have added an asset
#        When I set a grade for an asset
#        And Save it
#        Then this person shoukd receive this grade in that asset
#        # Grade is saved in /SSCN/AM_PERSON field AM_KN
#        And the grade should be shown in the grid
#
#    Scenario: Delete an asset qualification 
#        When I go to the Asset Group tab
#        And I have added an asset
#        When I select one or more assets
#        And I remove it
#        Then the asset should be removed from the grid
#        And the qualification grade for this person in this asset should be deleted from the table
#        # Table /SSCN/AM_PERSON