import datetime as dateTime
import pip
import xlrd
import scipy.stats
import math
import os
import xlsxwriter
import glob

#These two proeperties can be changed before running each interval length like 10sec / 227 sec or 300 sec 
# and interval count will also be changed accordingly 10sec==> 3240 , 227sec ==> 142 and 300sec ==> 108

intervalLength = 300 # 10 , 227
intervalCount = 108  #intervalCount is used whose value is calculated  as 8AM-5PM => 9 hrs * 60 * 60 / intervalLength

#xlsxwriter library is used to write the output excel file
resutWorkbook = xlsxwriter.Workbook("PyResult_Interval_"+str(intervalLength)+"Seconds.xlsx")
resultSheet1 = resutWorkbook.add_worksheet("Sheet1" )

#Data set files diretory path is hardcoded here and can be changed before execution if required.
#glob library is used to fetch the path where all the excel file will be read
filepath = "/Users/Dataset"
dirList = glob.glob(filepath+ "/*.xlsx")
sorted(dirList)

#Total time frame 8AM-5PM is divided into the no of count as intervalCount which will be further used as slots
#intervalCount is used to define the range 
timeIntervalList = []
startTime = dateTime.timedelta(0, 0, 0, 0, 0, 8)
timeIntervalList.append(startTime)
for i in range(intervalCount):
    time = startTime + dateTime.timedelta(seconds=intervalLength)
    timeIntervalList.append(time)


#List Property Initialization of octets , duration , realtime , wekk 1 and week 2 values and also for storing average values.
octets1 = []
octets2 = []
realTime1 = []
realTime2 = []
duration1 = []
duration2 = []
aWeek1 = []
bWeek1 = []
aWeek2 = []
bWeek2 = []
avrg1a = []
avrg2a = []
avrg2b = []
result = 0


#We are looping through aall the 54 files and since we will be comparing to users at a time we have to create 2 for loops so that we will cover all the 54*54 possible combinations of the users given
for x in range(0, 54):
    for y in range(0, 54):
        firstFileDir = dirList[x]
        secondFileDir = dirList[y]
        print("first user {} --> ".format(dirList[x]))
        print("second user {} ".format(dirList[y]))
        file1 = xlrd.open_workbook(firstFileDir)
        file2 = xlrd.open_workbook(secondFileDir)
        file1Sheet = file1.sheet_by_index(0)
        file2Sheet = file2.sheet_by_index(0)
        rows1 = file1Sheet.nrows
        rows2 = file2Sheet.nrows

        
        #this is the main data splitting code where we are eliminating the data which has a duration of 0 and which are weekendds and also the data whcih is not in the time frame from 8 to 5 pm.So basically we are taking the data which is in between 8am and 5 pm.
        for d in range(4, 16):
            
            if (d == 9 or d == 10):
                continue
            for i in range(1, rows1):

                octate = file1Sheet.cell_value(i, 3)
                duration = file1Sheet.cell_value(i, 9)
                if (duration == 0):
                    continue 
                realTime= file1Sheet.cell_value(i, 5)
                date1 = dateTime.datetime.fromtimestamp(realTime/ 1000).day
                if d != date1:
                    continue
                hour1 = dateTime.datetime.fromtimestamp(realTime / 1000).hour

                if (hour1 > 7 and hour1 < 18):
                    octets1.append(octate)
                    realTime1.append(realTime)
                    duration1.append(duration)

        #Doing the above mentioned work for second user in comparision
        for d in range(4, 16):
            
            if (d == 9 or d == 10):
                continue
            for i in range(1, rows2):

                octate = file2Sheet.cell_value(i, 3)
                duration = file2Sheet.cell_value(i, 9)
                if (duration == 0):
                    continue 
                realTime= file2Sheet.cell_value(i, 5)
                date2 = dateTime.datetime.fromtimestamp(realTime/ 1000).day
                if d != date2:
                    continue
                hour2 = dateTime.datetime.fromtimestamp(realTime / 1000).hour

                if (hour2 > 7 and hour2 < 18):
                    octets2.append(octate)
                    realTime2.append(realTime)
                    duration2.append(duration)


        def average(l):
              return sum(l) / len(l)

        
        for i in range(4, 9):
            
            for k in range(0, 108):
                for j in range(0, len(octets1)):
                    date = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).day
                    if (date != i):
                        continue

                    h1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).hour
                    s1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).second
                    m1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).minute
                    mi1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).microsecond
                    times = dateTime.timedelta(0, s1, mi1, 0, m1, h1)
                    if (times >= timeIntervalList[k] and times < timeIntervalList[k + 1]):
                        avrg1a.append(octets1[j] / duration1[j])
                if (len(avrg1a) == 0):
                    aWeek1.append(0)
                else:
                    aWeek1.append(average(avrg1a))
                avrg1a=[]
        avrg1a=[]
        
        for i in range(11, 16):
            for k in range(0, 108):
                for j in range(0, len(octets1)):
                    date = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).day
                    if (date != i):
                        continue

                    h1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).hour
                    s1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).second
                    m1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).minute
                    mi1 = dateTime.datetime.fromtimestamp(realTime1[j] / 1000).microsecond
                    times = dateTime.timedelta(0, s1, mi1, 0, m1, h1)
                    if (times >= timeIntervalList[k] and times < timeIntervalList[k + 1]):
                        avrg2a.append(octets1[j] / duration1[j])
                if (len(avrg2a) == 0):
                    aWeek2.append(0)
                else:
                    aWeek2.append(average(avrg2a))
                avrg2a=[]
        avrg2a=[]
    
        
        for i in range(11, 16):
            for k in range(0, 108):
                for j in range(0, len(octets2)):
                    date = dateTime.datetime.fromtimestamp(realTime2[j] / 1000).day
                    if (date != i):
                        continue

                    h1 = dateTime.datetime.fromtimestamp(realTime2[j] / 1000).hour
                    s1 = dateTime.datetime.fromtimestamp(realTime2[j] / 1000).second
                    m1 = dateTime.datetime.fromtimestamp(realTime2[j] / 1000).minute
                    mi1 = dateTime.datetime.fromtimestamp(realTime2[j] / 1000).microsecond
                    times = dateTime.timedelta(0, s1, mi1, 0, m1, h1)
                    if (times >= timeIntervalList[k] and times < timeIntervalList[k + 1]):
                        avrg2b.append(octets2[j] / duration2[j])
                
                if (len(avrg2b) == 0):
                    bWeek2.append(0)
                else:
                    bWeek2.append(average(avrg2b))
                avrg2b=[]
        avrg2b=[]

        
        #print(aWeek1)

        #Below is the calculation for the spearmann coefficient values and storing them respectively 
        sp1a2a = scipy.stats.spearmanr(aWeek1, aWeek2)[0]
        sp1a2b = scipy.stats.spearmanr(aWeek1, bWeek2)[0]
        sp2a2b = scipy.stats.spearmanr(aWeek2, bWeek2)[0]

        #In order to avoid divide by zero error for we are performing this function
        if (math.isnan(sp1a2a)):
            sp1a2a = 0.0
        elif (sp1a2a == 1):
            sp1a2a = 0.99

        if (math.isnan(sp1a2b)):
            sp1a2b = 0.0
        elif (sp1a2b == 1):
            sp1a2b = 0.99

        if (math.isnan(sp2a2b)):
            sp2a2b = 0.0
        elif (sp2a2b == 1):
            sp2a2b == 0.99        
        
        #Z value calculation
        rm2 = ((sp1a2a ** 2) + (sp1a2b ** 2)) / 2
        f = (1 - sp2a2b) / (2 * (1 - rm2))
        h = (1 - (f * rm2)) / (1 - rm2)
        z1a2a = 0.5 * (math.log10((1 + sp1a2a) / (1 - sp1a2a)))
        z1a2b = 0.5 * (math.log10((1 + sp1a2b) / (1 - sp1a2b)))
        if (sp2a2b == 1):
            sp2a2b = 0.99
        z = (z1a2a - z1a2b) * ((len(aWeek1) - 3) ** 0.5) / (2 * (1 - sp2a2b) * h)
        p = 0.3275911
        a1 = 0.254829592
        a2 = -0.284496736
        a3 = 1.421413741
        a4 = -1.453152027
        a5 = 1.061405429
        if z < 0.0:
            sign = -1
        else:
            sign = 1
        x1 = abs(z) / (2 ** 0.5)
        t = 1 / (1 + p * x1)
        erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x1 * x1)
        result = 0.5 * (1 + sign * erf)
        
        #Writing the results in the ouput excel
        
        if ( x==0 and y == 0):
          resultSheet1.write(0, 0, "Week 1 and 2")
        elif ( x==0 and y != 0):
          resultSheet1.write(0, y, "User"+str(y))
        elif ( y==0  and x != 0 ):
          resultSheet1.write(x, 0, "User"+str(x))
        else:
          resultSheet1.write(x+1 , y+1, result)


resutWorkbook.close()


