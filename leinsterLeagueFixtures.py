# FUNCTION 1 - returns a list of competitions and teams
def getTeamsAndComps(a):                    # takes excel entries table as input (i.e. 'a')
    import math
    import numpy as np
    import pandas as pd
    numRows = a.shape[0]
    numCols = a.shape[1]
    # print(numRows)
    # print(numCols)
    result = []
    compList = []
    teamList = []
    for i in range(1,numCols):              # for each column (i.e. competition)
        teams = []                          # each comp will have a unique set of teams
        comps = a[0][i]                     # get current competition
        result.append([comps])
        for j in range(1,numRows):
            if pd.isnull(a[j][i]):
                teams = teams               # if cell is blank, do nothing
            else:
                teams.append(a[j][0])       # otherwise, add the team
        result.append(teams)                # 'result' is storing comp, teams, comp, teams etc... alternating

    for i in range(0,len(result),2):        # comps start from 0 incrementing in 2
        compList.append(result[i])
    for i in range(1,len(result),2):        # teams start from 1 incrementing in 2
        teamList.append(result[i])
    return compList, teamList               # return each list

# FUNCTION 3
def getPlayMatrix(numTeams):
    import numpy as np
    if isOdd(numTeams):
        numCols = numTeams
    else:
        numCols = numTeams - 1
    numRows = (numTeams - numTeams%2)//2
    playMatrix = np.zeros((numRows,numCols),dtype='i,i')
    if isOdd(numTeams):
        matrixOdd(numRows,numCols,playMatrix)
    else:
        matrixEven(numRows,numCols,numTeams,playMatrix)
    return playMatrix

# FUNCTION 4
def isOdd(num):
    if num%2 == 1:
        return True
    else:
        return False

# FUNCTION 5
def matrixOdd(d,e,p):
    for i in range(d):
        for j in range(e):
            p[i][j][0]=j+1
            p[i][j][1]=j+1

# FUNCTION 6
def matrixEven(d,e,f,p):
    for i in range(d):
        for j in range(e):
            p[i][j][0]=j+1
            p[i][j][1]=j+1
    for k in range(e):
        p[d-1][k][0]=k+1
        p[d-1][k][1]=f

# FUNCTION 8 -
def getFinalMatrix(n,b,c):
    lx = [0 for i in range(b.shape[1])]
    ly = [0 for i in range(b.shape[1])]
    for i in range(len(c)):
        for j in range(b.shape[1]):
            lx[j] = b[i][j][0]
            ly[j] = b[i][j][1]
        mx = shift(lx,c[i][0])
        my = shift(ly,c[i][1])
        for j in range(b.shape[1]):
            b[i][j][0] = mx[j]
            b[i][j][1] = my[j]
    return(b)

# FUNCTION 9
def shift(l, n=0):
    a = n % len(l)
    return l[-a:] + l[:-a]

# def nCr(n,r):
#     import math
#     f = math.factorial
#     return f(n) / (f(r) * f(n-r))

# FUNCTION 7
def getShiftMatrix(b):
    numRows = b.shape[0]
    numCols = b.shape[1]
    if numCols <= 3:
        numRows_s = 1
    else:
        numRows_s = (numCols-1)//2
    numList = [i for i in range(1,numRows_s*2+1)]
    firstHalf = numList[:len(numList)//2]
    secondHalf = numList[len(numList)//2:]
    secondHalf2 = secondHalf[::-1] #reverse secondHalf
    cb = list(zip(firstHalf,secondHalf2))
    return cb

# FUNCTION 10
def getListOfPairs(x):
    l=[]
    for i in range(x.shape[0]):
        for j in range(x.shape[1]):
            l.append(x[i][j])
    return l

# FUNCTION 11
def shuffleList(l):
    import random
    return random.sample(l,len(l))

# FUNCTION 2 -
def writeToExcel(comps,teams):
    fixturesWb = '/Users/duncan/Documents/Committee-Leinster/Fixtures/21-22/LeagueFixtures21-22.xls'
    # fixturesWb = '/Users/duncan/Documents/Committee-Leinster/JnrKidsLeague2021/Outputs/KidsLeagueFixtures.xls'
    from xlwt import Workbook
    wb = Workbook()
    for i in range(len(comps)):
        # print(comps[i][0])
        # print('-')
        sheet1 = wb.add_sheet(comps[i][0]) #add new sheet for each competition
        a = len(teams[i])
        b = getPlayMatrix(a)                # FUNCTION 3,4,5,6
        print('b:',b)
        c = getShiftMatrix(b)               # FUNCTION 7
        print(c)
        x = getFinalMatrix(a,b,c)           # FUNCTION 8,9
        print(x)
        y = getListOfPairs(x)               # FUNCTION 10
        print(y)
        z1 = shuffleList(y)                 # FUNCTION 11
        z2 = shuffleList(y)                 # "
        z = z1 + z2 #double round-robin
        listArray=[]
        rcrdNo = 1
        for j in z:
            team1 = teams[i][j[0]-1]
            # print(team1,end=' - ')
            team2 = teams[i][j[1]-1]
            # print(team2)
            sheet1.write(rcrdNo-1, 0, rcrdNo)
            sheet1.write(rcrdNo-1, 1, team1)
            sheet1.write(rcrdNo-1, 2, 'vs.')
            sheet1.write(rcrdNo-1, 3, team2)
            rcrdNo += 1
        # print('')
    wb.save(fixturesWb)
    # print('Done!')

# import datetime
# now = datetime.datetime.now()
# dateObj = datetime.date(now)
# print(dateObj.strftime("%A"))

import pandas as pd
import numpy as np
entriesWb = '/Users/duncan/Documents/Committee-Leinster/Entries/LeinsterEntries21-22.xlsx'
# entriesWb = '/Users/duncan/Documents/Committee-Leinster/JnrKidsLeague2021/Inputs/JnrSummerLeagueEntries2021.xlsx'
df = pd.read_excel(entriesWb,header=None)
# print(df)
excelData = df.values
# print(excelData)
comps,teams = getTeamsAndComps(excelData)   # FUNCTION 1
# print(comps)
# print(teams)

writeToExcel(comps,teams)                   # FUNCTION 2 (... 3,4,5,6,7,8,9,10,11)
