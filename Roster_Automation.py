import pandas as pd
import datetime as dt
from datetime import datetime
import random
import openpyxl

filepath = "C:/temp/Roster/Input_Data.xlsx"

Resources = pd.read_excel(filepath, sheet_name='Resources')
Holidays = pd.read_excel(filepath, sheet_name='Holidays')

print("Enter Start date as instructed")
start_date = dt.date(int(input("Enter year. Eg; 2020 : ")), int(input("Enter Month. Eg: 1 for Jan : ")), int(input("Enter date. Eg: 1 to 31 : ")))
print("Enter End date as instructed")
end_date = dt.date(int(input("Enter year. Eg; 2020 : ")), int(input("Enter Month. Eg: 1 for Jan : ")), int(input("Enter date. Eg: 1 to 31 : ")))

roster=pd.DataFrame(columns=['Date',
                            'Week Number',
                             'Day',
                           'INDIA SHIFT 1 PRIMARY',
                           'INDIA SHIFT 1 Secondary',
                           'INDIA SHIFT 2 PRIMARY',
                           'INDIA SHIFT 2 Secondary',
                           'US SHIFT PRIMARY',
                           'US SHIFT SECONDARY',
                           'Weekend On-call Primary',
                           'Weekend On-call Secondary',
                           'Duty Manager 1',
                           'Duty Manager 2',
                           'Duty Manager Weekend'])

required_date = []
week_no = []
weekday = []

daterange = pd.date_range(start_date, end_date)
for single_date in daterange:
    required_date.append(single_date.strftime("%Y-%m-%d"))
    week_no.append(single_date.isocalendar()[1])
    weekday.append(single_date.strftime("%A"))

roster['Date'] = required_date
roster['Week Number'] = week_no
roster['Day'] = weekday

Holidays.India = Holidays.India.astype(str)
Holidays.US = Holidays.US.astype(str)

Holidays.US = Holidays.US.apply(lambda x : None if x=='NaT' else x)

US_holidays = list(filter(None,list(Holidays.US)))
IND_holidays = list(Holidays.India)

# segregate holidays whether common holiday or region wise
def Optimize_IND_Holidays(IND_holidays):
    for i in IND_holidays:
        if(i < roster['Date'][0]):
            IND_holidays.remove(i)
            return Optimize_IND_Holidays(IND_holidays)
Optimize_IND_Holidays(IND_holidays)

def Optimize_US_Holidays(US_holidays):
    for i in US_holidays:
        if(i < roster['Date'][0]):
            US_holidays.remove(i)
            return Optimize_US_Holidays(US_holidays)
Optimize_US_Holidays(US_holidays)

common_holidays = []
for i in IND_holidays:
    for j in US_holidays:
        if i == j:
            common_holidays.append(i)
            IND_holidays.remove(i)
            US_holidays.remove(i)

common_holidays,IND_holidays , US_holidays

# segregate resources based on region
IND_Res = list(Resources.IND)
US_Res = list(Resources.US)
US_Res = [x for x in US_Res if type(x) is not float]
Duty_Manager = list(Resources.DutyManager)
Duty_Manager = [x for x in Duty_Manager if type(x) is not float]
TeamL = list(Resources.TL)
TeamL = [x for x in TeamL if type(x) is not float]
US_Res, IND_Res, Duty_Manager, TeamL

# get week numbers as list
weekno = list(roster['Week Number'])
shift1 = {}
shift2 = {}
shift3 = {}
# make a set from the week numbers list and sort them to iterate
for i in sorted(list(set(weekno))):
    k = i%len(IND_Res)
    m = i%len(US_Res)
    shift1.update({i:[IND_Res[k],IND_Res[k-1]]})
    shift2.update({i:[IND_Res[k-2],IND_Res[k-3]]})
    shift3.update({i:[US_Res[m],US_Res[m-1]]})

weekend_Resources = IND_Res+US_Res
weekend_Managers = Duty_Manager+TeamL
weekend_combinations = [[a, b] for a in weekend_Resources  
          for b in weekend_Managers if a != b]
		  
for i in weekend_Resources:
    if i in weekend_Managers:
        weekend_Resources.remove(i)
       
IND_shift1_prim = []
IND_shift1_sec = []
IND_shift2_prim = []
IND_shift2_sec = []
US_shift_prim = []
US_shift_sec = []
Duty1 = []
Duty2 = []
WeekendDutyMgr = []
Weekend_prim = []
Weekend_sec = []
for j in sorted(list(set(weekno))):
    k = j%len(weekend_Resources)
    m = j%len(weekend_Managers)
    for i in range(len(roster)):
        if((roster['Week Number'][i]==j) and ((roster['Day'][i] != 'Saturday') and (roster['Day'][i] != 'Sunday'))):
            IND_shift1_prim.append(shift1[j][0])
            IND_shift1_sec.append(shift1[j][1])
            IND_shift2_prim.append(shift2[j][0])
            IND_shift2_sec.append(shift2[j][1])
            US_shift_prim.append(shift3[j][0])
            US_shift_sec.append(shift3[j][1])
            Duty1.append(Resources.DutyManager[0])
            Duty2.append(Resources.DutyManager[1])
            Weekend_prim.append("")
            Weekend_sec.append("")
            WeekendDutyMgr.append("")
        if((roster['Week Number'][i]==j) and ((roster['Day'][i] == 'Saturday') or (roster['Day'][i] == 'Sunday'))):
            IND_shift1_prim.append("")
            IND_shift1_sec.append("")
            IND_shift2_prim.append("")
            IND_shift2_sec.append("")
            US_shift_prim.append("")
            US_shift_sec.append("")
            Duty1.append("")
            Duty2.append("")
            Weekend_prim.append("")
            Weekend_sec.append("")
            WeekendDutyMgr.append("")

# For weekends
saturday_lst = []
for i in range(len(roster)):
    if roster['Day'][i] == 'Saturday':
        saturday_lst.append(roster['Date'][i])

weekend_shift = {}
for i in range(len(saturday_lst)):
    k = i%len(weekend_Resources)
    m = i%len(weekend_Managers)
    weekend_shift.update({saturday_lst[i]:[weekend_Resources[k],weekend_Managers[m]]})

weekend_shift

for i in range(len(roster)):
    if roster['Date'][i] in weekend_shift.keys():
            Weekend_prim[i] = weekend_shift[roster['Date'][i]][0]
            Weekend_sec[i] = weekend_shift[roster['Date'][i]][1]
            Weekend_prim[i+1] = weekend_shift[roster['Date'][i]][1]
            Weekend_sec[i+1] = weekend_shift[roster['Date'][i]][0]
            WeekendDutyMgr[i] = weekend_shift[roster['Date'][i]][1]
            WeekendDutyMgr[i+1] = weekend_shift[roster['Date'][i]][1]         

# For Holidays
com_hol_shift = {}
IND_hol_shift = {}
US_hol_shift = {}
for i in range(len(common_holidays)):
    m = i%len(US_Res)
    k = i%len(IND_Res)
    com_hol_shift.update({common_holidays[i]:[US_Res[m],IND_Res[m]]})
for i in range(len(IND_holidays)):
    m = i%len(US_Res)
    IND_hol_shift.update({IND_holidays[i]: [US_Res[m],US_Res[m-(len(US_Res)//2)]]})
for i in range(len(US_holidays)):
    m = i%len(IND_Res)
    US_hol_shift.update({US_holidays[i]:[IND_Res[m],IND_Res[m-(len(IND_Res)//2)]]})
#com_hol_shift, IND_hol_shift, US_hol_shift
# For common holidays
for i in range(len(roster)):
    if roster['Date'][i] in com_hol_shift.keys():
        IND_shift1_prim[i] = com_hol_shift[roster['Date'][i]][0]
        IND_shift1_sec[i] = com_hol_shift[roster['Date'][i]][1]
        IND_shift2_prim[i] = com_hol_shift[roster['Date'][i]][1]
        IND_shift2_sec[i] = com_hol_shift[roster['Date'][i]][0]
        US_shift_prim[i] = com_hol_shift[roster['Date'][i]][0]
        US_shift_sec[i] = com_hol_shift[roster['Date'][i]][1]
        Duty1[i] = Resources.DutyManager[0]
        Duty2[i] = Resources.DutyManager[1]

# For India holidays
for i in range(len(roster)):
    if roster['Date'][i] in IND_hol_shift.keys():
        IND_shift1_prim[i] = IND_hol_shift[roster['Date'][i]][0]
        IND_shift1_sec[i] = IND_hol_shift[roster['Date'][i]][1]
        IND_shift2_prim[i] = IND_hol_shift[roster['Date'][i]][1]
        IND_shift2_sec[i] = IND_hol_shift[roster['Date'][i]][0]
        US_shift_prim[i] = IND_hol_shift[roster['Date'][i]][0]
        US_shift_sec[i] = IND_hol_shift[roster['Date'][i]][1]
        Duty1[i] = Resources.DutyManager[0]
        Duty2[i] = Resources.DutyManager[0]
# For US holidays        
for i in range(len(roster)):
    if roster['Date'][i] in US_hol_shift.keys():
        IND_shift1_prim[i] = IND_shift1_prim[i] #US_hol_shift[roster['Date'][i]][0]
        IND_shift1_sec[i] = IND_shift1_sec[i] #US_hol_shift[roster['Date'][i]][1]
        IND_shift2_prim[i] = IND_shift2_prim[i] #US_hol_shift[roster['Date'][i]][1]
        IND_shift2_sec[i] = IND_shift2_sec[i] #US_hol_shift[roster['Date'][i]][0]
        US_shift_prim[i] = US_hol_shift[roster['Date'][i]][0]
        US_shift_sec[i] = US_hol_shift[roster['Date'][i]][1]
        Duty1[i] = Resources.DutyManager[1]
        Duty2[i] = Resources.DutyManager[1]        

# make sure length of each list is same as number of total days for which roster is being prepared
len(IND_shift1_prim),len(IND_shift1_sec),len(IND_shift2_prim),len(IND_shift2_sec),len(US_shift_prim),len(US_shift_sec),len(Duty1),len(Duty2),len(roster)

roster['INDIA SHIFT 1 PRIMARY'] = IND_shift1_prim
roster['INDIA SHIFT 1 Secondary'] = IND_shift1_sec
roster['INDIA SHIFT 2 PRIMARY'] = IND_shift2_prim
roster['INDIA SHIFT 2 Secondary'] = IND_shift2_sec
roster['US SHIFT PRIMARY'] = US_shift_prim
roster['US SHIFT SECONDARY'] = US_shift_sec
roster['Duty Manager 1'] = Duty1
roster['Duty Manager 2'] = Duty2
roster['Weekend On-call Primary'] = Weekend_prim
roster['Weekend On-call Secondary'] = Weekend_sec
roster['Duty Manager Weekend'] = WeekendDutyMgr

# Select Monthly roster
Jan_roster = roster[(roster['Date']>='2020-01-01') & (roster['Date']<='2020-01-31')]
Feb_roster = roster[(roster['Date']>='2020-02-01') & (roster['Date']<='2020-02-29')]
Mar_roster = roster[(roster['Date']>='2020-03-01') & (roster['Date']<='2020-03-31')]
Apr_roster = roster[(roster['Date']>='2020-04-01') & (roster['Date']<='2020-04-30')]
May_roster = roster[(roster['Date']>='2020-05-01') & (roster['Date']<='2020-05-31')]
June_roster = roster[(roster['Date']>='2020-06-01') & (roster['Date']<='2020-06-30')]
July_roster = roster[(roster['Date']>='2020-07-01') & (roster['Date']<='2020-07-31')]
August_roster = roster[(roster['Date']>='2020-08-01') & (roster['Date']<='2020-08-31')]
Sep_roster = roster[(roster['Date']>='2020-09-01') & (roster['Date']<='2020-09-30')]
Oct_roster = roster[(roster['Date']>='2020-10-01') & (roster['Date']<='2020-10-31')]
Nov_roster = roster[(roster['Date']>='2020-11-01') & (roster['Date']<='2020-11-30')]
Dec_roster = roster[(roster['Date']>='2020-12-01') & (roster['Date']<='2020-12-31')]

with pd.ExcelWriter('roster.xlsx') as writer:
    Jan_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'January')
    Feb_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'February')
    Mar_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'March')
    Apr_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'April')
    May_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'May')
    June_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'June')
    July_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'July')
    August_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'August')
    Sep_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'September')
    Oct_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'October')
    Nov_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'November')
    Dec_roster.drop(['Week Number'],axis=1).T.to_excel(writer, sheet_name = 'December')

'''
wb = openpyxl.load_workbook(r'C:\temp\Roster\roster.xlsx')
for i in range(len(wb.worksheets)):
    ws = wb.worksheets[i]
    img = openpyxl.drawing.image.Image(r'C:\temp\Roster\oncall-schedule.PNG')
    img.anchor = 'G18'
    ws.add_image(img)
    ws.delete_rows(1)
wb.save(r'C:\temp\Roster\roster.xlsx')
'''