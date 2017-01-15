import nba_py
import openpyxl
from openpyxl import load_workbook
import os
from datetime import date
os.chdir('c:\\PythonScripts\\Excel')

wb = load_workbook('Team_Stats_Raw.xlsx')



from nba_py.constants import CURRENT_SEASON
#print(CURRENT_SEASON)

from nba_py import constants
#print(constants.SeasonType.Regular)

from nba_py import game



def getStats(gameList):
    for a in gameList:
        boxscore_summary = game.BoxscoreSummary(a)
        sql_team_basic = boxscore_summary.game_summary()

        boxscore_advanced = game.BoxscoreAdvanced(a)
        sql_team_advanced = boxscore_advanced.sql_team_advanced()

        
        for i in range(0,2):
            GAMEID = sql_team_advanced[i]['GAME_ID']
            DATE = sql_team_basic[0]['GAME_DATE_EST']
            GAMECODE = sql_team_basic[0]['GAMECODE']
            TEAMID = sql_team_advanced[i]['TEAM_ID']
            HOMETEAMID = sql_team_basic[0]['HOME_TEAM_ID']
            AWAYTEAMID = sql_team_basic[0]['VISITOR_TEAM_ID']
            TEAM = sql_team_advanced[i]['TEAM_ABBREVIATION']
            ORTG = sql_team_advanced[i]['OFF_RATING']
            DRTG = sql_team_advanced[i]['DEF_RATING']
            NET =  sql_team_advanced[i]['NET_RATING']
            Pace = sql_team_advanced[i]['PACE']
            PIE = sql_team_advanced[i]['PIE']

            if HOMETEAMID == TEAMID:
                LOCATION = 'Home'
            else:
                LOCATION = 'Away'

            DATE = DATE[0:10]
            
            statsList = []
            statsList.append(GAMEID)
            statsList.append(DATE)
            statsList.append(GAMECODE)
            statsList.append (TEAMID)
            statsList.append(HOMETEAMID)
            statsList.append(AWAYTEAMID)
            statsList.append (TEAM)
            statsList.append (LOCATION)
            statsList.append (ORTG)
            statsList.append (DRTG)
            statsList.append (Pace)
            statsList.append (PIE)

            sheet = wb.get_sheet_by_name(str(TEAM))
            row_cell = sheet.max_row + 1
            column_cell = 1
            for n in range(len(statsList)):
                print(statsList[n])
                sheet.cell(row = row_cell, column = column_cell).value = statsList[n]
                column_cell= column_cell + 1
    wb.save('Team_Stats_Raw.xlsx')

    

def formatAllDates():
    tabs = wb.get_sheet_names()
    for j in tabs:
        tab = wb.get_sheet_by_name(j)
        row_cell = 2
        column_cell = 2
        row_count = tab.max_row
        for i in range(row_count-1):
            date = tab.cell(row = row_cell, column = column_cell).value
            date = date[0:10]
            tab.cell(row = row_cell, column = column_cell).value = date
            row_cell = row_cell + 1

    wb.save('Team_Stats_Raw.xlsx')

#Move all to Master
def moveAlltoMaster():
    tabs = wb.get_sheet_names()
    for j in tabs:
        if j != 'Master' and j!= 'ORTG' and j != 'DRTG' and j!= 'Pace':
            tab = wb.get_sheet_by_name(j)
            tab_row_count = tab.max_row
            tab_column_count = tab.max_column
            tab_row_cell = 2
            for i in range(tab_row_count-1):             
                tab_column_cell = 1
                values = []
                for x in range(tab_column_count):
                    data = tab.cell(row = tab_row_cell, column = tab_column_cell).value
                    values.append(data)
                    tab_column_cell = tab_column_cell +1
                master = wb.get_sheet_by_name('Master')
                master_row_cell = master.max_row + 1
                master_column_cell = 1
                for m in range(len(values)):
                    master.cell(row = master_row_cell, column = master_column_cell).value = values[m]
                    master_column_cell= master_column_cell + 1   
                tab_row_cell = tab_row_cell + 1
    wb.save('Team_Stats_Raw.xlsx')


def calcAllRest():    
    tabs = wb.get_sheet_names()
    for j in tabs: 
        if j!= 'ORTG' and j != 'DRTG' and j!= 'Pace':
            tab = wb.get_sheet_by_name(j)
            row_cell = 2
            row_count = tab.max_row
            rest_column_cell = 13
            date_column_cell = 2
            for i in range(row_count-1):
                compare_row_cell = row_cell - 1
                rest = 0
                if row_cell == 2:
                    tab.cell(row = row_cell, column = rest_column_cell).value = 3
                else:
                    d0 = tab.cell(row = row_cell, column = date_column_cell).value
                    d1 = tab.cell(row = compare_row_cell, column = date_column_cell).value
                    d0 = date(int(d0[0:4]),int(d0[5:7]),int(d0[8:10]))
                    d1 = date(int(d1[0:4]),int(d1[5:7]),int(d1[8:10]))
                    print(d0,d1)
                    rest = d0-d1
                    tab.cell(row = row_cell, column = rest_column_cell).value = rest.days - 1
                row_cell = row_cell + 1
    wb.save('Team_Stats_Raw.xlsx')

def calcRest():
    tabs = wb.get_sheet_names()
    for j in tabs:
        if j!= 'ORTG' and j != 'DRTG' and j!= 'Pace':
            tab = wb.get_sheet_by_name(j)
            row_cell = tab.max_row
            rest_column_cell = 13
            date_column_cell = 2
            if tab.cell(row = row_cell, column = rest_column_cell).value == None:
                compare_row_cell = row_cell - 1
                d0 = tab.cell(row = row_cell, column = date_column_cell).value
                d1 = tab.cell(row = compare_row_cell, column = date_column_cell).value
                d0 = date(int(d0[0:4]),int(d0[5:7]),int(d0[8:10]))
                d1 = date(int(d1[0:4]),int(d1[5:7]),int(d1[8:10]))
                rest = d0-d1
                tab.cell(row = row_cell, column = rest_column_cell).value = rest.days - 1
            row_cell = row_cell + 1
    wb.save('Team_Stats_Raw.xlsx')

        
def addColumn():#Use to add a new column to tabs
    tabs = wb.get_sheet_names()
    for j in tabs: 
        tab = wb.get_sheet_by_name(j)
        tab.cell(row = 1, column = 14).value = 'Rest_Detail'
    wb.save('Team_Stats_Raw.xlsx')



def getAllRestDetail():
    tabs = wb.get_sheet_names()
    for j in tabs:
        if  j!= 'ORTG' and j != 'DRTG' and j!= 'Pace':
            tab = wb.get_sheet_by_name(j)
            row_cell = 2
            row_count = tab.max_row
            rest_column_cell = 13
            rest_detail_column = 14
            for i in range(row_count-1):
                r1 = row_cell - 1
                r2 = row_cell - 2
                r3 = row_cell - 3
                rest1 = 0
                rest2 = 0
                rest3 = 0
                if row_cell >2:
                    rest1 = tab.cell(row = r1, column = rest_column_cell).value
                if row_cell >3:
                    rest2 = tab.cell(row = r2, column = rest_column_cell).value
                if row_cell >4:
                    rest3 = tab.cell(row = r3, column = rest_column_cell).value
                restToday = tab.cell(row = row_cell, column = rest_column_cell).value
                threeInFour = restToday + rest1
                fourInFive = threeInFour + rest2
                print(rest1)
                if threeInFour <=1:
                    tab.cell(row = row_cell, column = rest_detail_column).value = '3IN4'
                if fourInFive <=1 :
                    tab.cell(row = row_cell, column = rest_detail_column).value = '4IN5'
                elif threeInFour > 1 and fourInFive >1:
                    tab.cell(row = row_cell, column = rest_detail_column).value = 'None'
                row_cell = row_cell + 1
    wb.save('Team_Stats_Raw.xlsx')

def getRestDetail():
    tabs = wb.get_sheet_names()
    for j in tabs:
        if j != 'Master' and j!= 'ORTG' and j != 'DRTG' and j!= 'Pace':
            tab = wb.get_sheet_by_name(j)
            row_cell = tab.max_row
            rest_column_cell = 13
            rest_detail_column = 14
            if tab.cell(row = row_cell, column = rest_detail_column).value == None:
                r1 = row_cell - 1
                r2 = row_cell - 2
                r3 = row_cell - 3
                rest1 = 0
                rest2 = 0
                rest3 = 0
                if row_cell >2:
                    rest1 = tab.cell(row = r1, column = rest_column_cell).value
                if row_cell >3:
                    rest2 = tab.cell(row = r2, column = rest_column_cell).value
                if row_cell >4:
                    rest3 = tab.cell(row = r3, column = rest_column_cell).value
                restToday = tab.cell(row = row_cell, column = rest_column_cell).value
                threeInFour = restToday + rest1
                fourInFive = threeInFour + rest2
                print(rest1)
                if threeInFour <=1:
                    tab.cell(row = row_cell, column = rest_detail_column).value = '3IN4'
                if fourInFive <=1 :
                    tab.cell(row = row_cell, column = rest_detail_column).value = '4IN5'
                elif threeInFour > 1 and fourInFive >1:
                    tab.cell(row = row_cell, column = rest_detail_column).value = 'None'
            row_cell = row_cell + 1
    wb.save('Team_Stats_Raw.xlsx')



