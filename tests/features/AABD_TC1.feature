Feature: Annual Availability by Division

Scenario: Normal case with employees assigned across divisions

Given dataset "AABD_TC1.csv" 

When the dataset is injected into "AABD.xlsm" 
And all charts are visualized

Then "Charts.ShowYearByDivision" should correctly display the number of employees per division
 

