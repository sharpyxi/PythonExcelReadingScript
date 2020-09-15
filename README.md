# Excel reading script

This(test.py) is a user friendly Python script that reads given Excel spreadsheet "Site_Capacity" and returns sum of skill for given date. 

In order for script to work The Excel sheet "Site_Capacity" must be in same directory as the script and the sheet with the data(in this case "Sheet1") must be the active sheet.

It takes 3 arguments: 

Skill
  - all available Skills in the table are diplayed to user.
  
  Year
  - number.
  
Month
  - number between 1 and 12.
  
The script validates the inputs and tells user when the input is invalid and why and checks if given inputs are in the Excel table.
If the user chooses a date that is not in the table, the script shows all available dates in the table.

Since the script shows all the vital data and validates input, the user should have no need to open the Excel file.

As long as the columns: Location, Skill and Member do not change, the script should supports adding new rows of data as well as new date columns.

After execution it starts again at the begining so the user can calculate another sum.






