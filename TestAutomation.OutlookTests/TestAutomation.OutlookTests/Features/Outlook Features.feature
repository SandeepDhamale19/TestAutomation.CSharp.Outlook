Feature: Outlook Features
	In order to verify outlook features
	As an outlook user
	I want to be told the sum of two numbers

@Unit
Scenario: Verify outlook email by Subject and Body
	Given I send an email to Captain.America@hotmail.com
	Then I recieve an email in my inbox

@Unit
Scenario: Sort outlook emails by recieved date
	Given I send an email to Captain.America@hotmail.com
	Then I can get email items sorted by recieved date

@Unit
Scenario: Get and Save Attachemnt from outllok mail
	Given I send an email to Captain.America@hotmail.com
	Then I can verify email has attahcments
	And I can get attachment from mail
	And I can save attachment from mail

@Unit
Scenario: Read outlook email
	Given I send an email to Captain.America@hotmail.com
	Then I can read the recived email

@Unit
Scenario: Read saved email
	Given I have saved email
	Then I can read saved email

@Unit
Scenario: Read outlook email body line by line
	Given I send an email to Captain.America@hotmail.com
	Then I can read the recived email body line by line

@browser
Scenario: Click on outlook email link
	Given I send an email to Captain.America@hotmail.com
	Then I can click on link in the recived email