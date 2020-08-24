
import datetime
import pandas as pd
from pandas import ExcelFile

def script():
    df = pd.read_excel('Site_Capacity.xlsx')

    availableSkills = df.Skill.unique()

   
    inputSkill = input("Select skill " + str(availableSkills) + " ").upper()
    inputYear = int(input("Select Year "))
    inputMonth = int(input("Select Month 1-12 "))

    selectedDate = datetime.datetime(inputYear, inputMonth, 1)

    filterSkill = df['Skill'] == inputSkill
    df1 = df[filterSkill]
    df2 = df1[selectedDate].head()


    print(df2.sum(axis=0, skipna = True))

while True:
    
    script()


