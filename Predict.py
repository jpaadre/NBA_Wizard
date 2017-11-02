import openpyxl
from openpyxl import load_workbook
import os
from datetime import date
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from Stat_Finder import getGames, getToday,getYesterday,getADVStats,getTodaysGames
from Format_data import CalcStats,getDataSet, getDataSetcsv
from NBA_RF_Regression import LoadModel
os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')

year = '2017'

today = getToday()
yesterday = getYesterday()
print(today)

# -----------Get Yesterdays Games and Add to existing game Data--------------
def GetYesterdaysData():
    yesterdaysGames = getGames(yesterday)
    df = getADVStats(yesterdaysGames)

    writer = ExcelWriter('yesterdaysGames.xlsx')
    df.to_excel(writer,'Master')
    writer.save()
    df1 = getDataSet('AllStats_' + year + '.xlsx')
    df2 = getDataSet('yesterdaysGames.xlsx')
    frames = [df1,df2]
    print(df1.head(),df2.head())
    df3= pd.concat(frames)

    writer1 = ExcelWriter('AllStats_' + year + '.xlsx')
    df3.to_excel(writer1,'Master')
    writer1.save()



def GetTodaysData():
    todaysGames = getTodaysGames(today)
    game = []

    df = getDataSet('DataForModel_'+ year +'.xlsx')
    df = df.dropna()

    for i in todaysGames:
        hometeam = i[12:]
        awayteam = i[9:12]
        game.append([hometeam,awayteam])
    print(game)

    df = getDataSet('DataForModel_'+ year +'.xlsx')
    df = df.dropna()
    dfb = pd.DataFrame(columns=['Match','TEAM_ABBREVIATION_x_x','HomeIndex_x','AvgPace_x_x','std_AvgORTG_x_x','HomeORTG_x_x','std_AvgORTG_L5_x_x',
    'AvgDRTG_x_x','DaysRest_x','TEAM_ABBREVIATION_x_y','HomeIndex_x','AvgPace_x_y','std_AvgORTG_x_y','HomeORTG_x_y','std_AvgORTG_L5_x_y','AvgDRTG_x_y',
    'DaysRest_y'])
    for x in game:
        match = x[0] + x[1]
        df1 = df.loc[df['TEAM_ABBREVIATION_x'] == x[0]]
        df2 = df.loc[df['TEAM_ABBREVIATION_x'] == x[1]]
        df1 = df1.tail(1)
        df2 = df2.tail(1)
        df1 = df1[['GAMECODE','TEAM_ABBREVIATION_x','AvgPace_x','std_AvgORTG_x','HomeORTG_x','std_AvgORTG_L5_x','AvgDRTG_x']]
        df2 = df2[['GAMECODE','TEAM_ABBREVIATION_x','AvgPace_x','std_AvgORTG_x','HomeORTG_x','std_AvgORTG_L5_x','AvgDRTG_x']]
        df1['Match'] = match
        df2['Match'] = match
        df1['HomeIndex'] = 0
        df2['HomeIndex'] = 1
        df1['GAME_DATE'] = pd.to_datetime(df1['GAMECODE'].str[:9])
        df2['GAME_DATE'] = pd.to_datetime(df2['GAMECODE'].str[:9])
        df1['today'] = pd.to_datetime(today)
        df2['today'] = pd.to_datetime(today)
        df1['DaysRest'] = (df1['today'] - df1['GAME_DATE']).astype('timedelta64[D]')
        df2['DaysRest'] = (df2['today'] - df2['GAME_DATE']).astype('timedelta64[D]')
        # df1['DaysRest'] = df1['DaysRest'].days
        df3 = pd.merge(df1, df2, on='Match',how='outer')
        df4 = pd.merge(df2, df1, on='Match',how='outer')
        df3 = df3[['Match','TEAM_ABBREVIATION_x_x','HomeIndex_x','AvgPace_x_x','std_AvgORTG_x_x','HomeORTG_x_x','std_AvgORTG_L5_x_x',
        'AvgDRTG_x_x','DaysRest_x','TEAM_ABBREVIATION_x_y','HomeIndex_x','AvgPace_x_y','std_AvgORTG_x_y','HomeORTG_x_y','std_AvgORTG_L5_x_y','AvgDRTG_x_y',
        'DaysRest_y']]
        df4 = df4[['Match','TEAM_ABBREVIATION_x_x','HomeIndex_x','AvgPace_x_x','std_AvgORTG_x_x','HomeORTG_x_x','std_AvgORTG_L5_x_x',
        'AvgDRTG_x_x','DaysRest_x','TEAM_ABBREVIATION_x_y','HomeIndex_x','AvgPace_x_y','std_AvgORTG_x_y','HomeORTG_x_y','std_AvgORTG_L5_x_y','AvgDRTG_x_y',
        'DaysRest_y']]
        dfb = dfb.append(df3)
        dfb = dfb.append(df4)
        # print(df3.head(1))
        # print(df4.head(1))
    dfb = dfb.dropna()
    dfb['ProjectedPace'] = (dfb['AvgPace_x_x'] + dfb['AvgPace_x_y'])/2
    dfb = dfb[['TEAM_ABBREVIATION_x_x','ProjectedPace','DaysRest_x','AvgPace_x_x','std_AvgORTG_x_x','HomeORTG_x_x','std_AvgORTG_L5_x_x','AvgDRTG_x_y']]
    writer1 = ExcelWriter('TempToday.xlsx')
    dfb.to_excel(writer1,'Master')
    writer1.save()
    return dfb




def RunModelOnToday(df,odds):
    loaded_model = LoadModel()
    x = df[['DaysRest_x','AvgPace_x_x','std_AvgORTG_x_x','HomeORTG_x_x','std_AvgORTG_L5_x_x','AvgDRTG_x_y']].values
    # x[0] = x[0].total_days
    # print(x)
    pred = loaded_model.predict(x)
    # print(pred)
    df1 = pd.DataFrame({'TEAM_ABBREVIATION_x_x':df['TEAM_ABBREVIATION_x_x'],'Predicted':pred})
    df2 = pd.merge(df, df1, on='TEAM_ABBREVIATION_x_x',how='outer')
    df2['ProjectedScore'] = (df2['ProjectedPace'] * df2['Predicted'])/100
    df3 = pd.merge(df2, odds, on='TEAM_ABBREVIATION_x_x',how='outer')
    df3['CalculatedSpread'] = df3['ProjectedScore'].shift(1) - df3['ProjectedScore']
    df3['CalculatedOU'] = df3['ProjectedScore'].shift(-1) + df3['ProjectedScore']
    df3['CalculatedLines'] = df3.apply(getCalcedLines,axis=1)
    df3 = df3[['GAMECODE_x','TEAM_ABBREVIATION_x_x','ProjectedScore','CalculatedLines','VegasLines']]
    df3['Difference'] = df3['CalculatedLines'] - df3['VegasLines']
    df3['BetGrade'] = df3.apply(gradeBet,axis=1)
    writer1 = ExcelWriter('Projections.xlsx')
    df3.to_excel(writer1,'Master')
    writer1.save()
    print (df3.head())

def getCalcedLines(df):
    if df['VegasLines'] > 100:
        return df['CalculatedOU']
    else:
        return df['CalculatedSpread']

def gradeBet(df):
    if df['Difference'] >= 8 or df['Difference'] <= -8:
        return 'A'
    elif (df['Difference'] >= 4 and df['Difference'] <= 4) or (df['Difference'] <= -4 and df['Difference'] >= -8)  :
        return 'B'
    elif (df['Difference'] > 2 and df['Difference'] < 4) or (df['Difference'] < -2 and df['Difference'] > -4)  :
        return 'C'
    else:
        return 'D'


def GetOdds():
    df1 = getDataSet('Season_Odds.xlsx')
    df = getDataSetcsv('SBR_NBA_Lines.csv')
    df['AwayTeam'] = df.apply(getTeams,axis=1)
    df['HomeTeam'] = df.apply(getOppTeams,axis=1)
    df['key'] = df['key'].astype(str)
    df['GAMECODE_x'] = df.apply(getGameCode,axis=1)
    df['VegasLines'] = df.apply(filterOdds,axis=1)
    df['TEAM_ABBREVIATION_x_x'] = df['AwayTeam']
    writer1 = ExcelWriter('Todays_Odds.xlsx')
    df = df[['GAMECODE_x','TEAM_ABBREVIATION_x_x','VegasLines']]
    df.to_excel(writer1,'Master')

    df2 = df1.append(df)
    writer2 = ExcelWriter('Season_Odds.xlsx')
    df2.to_excel(writer2,'Master')
    writer1.save()
    return df

def getTeams(df):
    return EditTeamList[df['team']]

def getOppTeams(df):
    return EditTeamList[df['opp_team']]

def filterOdds(df):
    if df['rl_time'] == 'home':
        return (df['tot_PIN_line'])
    else:
        return (df['rl_PIN_line'])

def getGameCode(df):
    if df['rl_time'] == 'away':
        return(df['key'] + '/'+ df['AwayTeam'] + df['HomeTeam'])
    else:
        return(df['key'] + '/'+ df['HomeTeam'] + df['AwayTeam'])


EditTeamList = {'Atlan':'ATL','Bost':'BOS','Brookl':'BKN','Charlot':'CHA','Chica':'CHI','Clevela':'CLE','Dall':'DAL','Denv':'DEN','Detro':'DET',
        'Golden Sta':'GSW','Houst':'HOU','India':'IND','L.A. Clippe':'LAC','L.A. Lake':'LAL','Memph':'MEM','Mia':'MIA','Milwauk':'MIL',
        'Minneso':'MIN','New Orlea':'NOP','New Yo':'NYK','Oklahoma Ci':'OKC','Orlan':'ORL','Philadelph':'PHI','Phoen':'PHX','Portla':'POR',
        'Sacramen':'SAC','San Anton':'SAS','Toron':'TOR','Ut':'UTA','Washingt':'WAS'}

# --------------Calculate all Stats on 2017 Data------------------
# GetYesterdaysData()
# # --------------Calculate all Stats on 2017 Data------------------
# CalcStats(year)

odds = GetOdds()
df = GetTodaysData()
RunModelOnToday(df,odds)
