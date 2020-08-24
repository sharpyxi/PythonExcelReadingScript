
import datetime
import pandas as pd
from pandas import ExcelFile

def script():

    #loads Excel file into dataframe
    df = pd.read_excel('Site_Capacity.xlsx')

    #Checks available skills and dates in Excel table
    availableSkills = df.Skill.unique()
    availableDates = df.columns[3 : df.columns.size]

    
    #Ask for Skill input until its valid
    while True:
        inputSkill = input("Select skill " + str(availableSkills) + " ").upper()
        if inputSkill not in df.Skill.unique():
            print("Choose available Skill form the list")
        else:
            break

    # Asks for Year and Month input until valid and check if date exists in table        
    while True:
        global inputYear
        global inputMonth
        global selectedDate
        try:
            inputYear = int(input("Select Year "))
            
        except ValueError:
            print("Year must be a number")
            continue

        if inputYear<0:
            print("enter positive number")
            continue

        try:
            inputMonth = int(input("Select Month 1-12 "))

        except ValueError:
            print("Month must be a number")
            continue

        if inputMonth<0:
            print("enter positive number")
            continue

        if inputMonth > 12:
            print("Month is between 1 and 12")
            continue    

        selectedDate = datetime.datetime(inputYear, inputMonth, 1)  

        if selectedDate not in availableDates:
            print("Entered date is not in the table. Available dates are: " + str(pd.to_datetime(availableDates).strftime("%Y %m"))[7:-24])
        else:
            break

    
    # Filters dataframe rows by selected skill and by colums for selected date        
    filterSkill = df['Skill'] == inputSkill
    df1 = df[filterSkill]
    df2 = df1[selectedDate].head()

    # Prints sum of skill for selected date
    print(df2.sum(axis=0, skipna = True))

#Repeats process
while True:
    script()

