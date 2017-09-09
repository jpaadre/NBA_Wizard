import openpyxl
from openpyxl import load_workbook
import os
from datetime import date
import pandas as pd
from pandas import ExcelWriter
os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')

WORKING_DATA_FILE = 'AllStats_2016_Split.xlsx'

wb = load_workbook(WORKING_DATA_FILE)

teamList = ['ATL','BOS','BKN','CHA','CHI','CLE','DAL','DEN','DET','GSW','HOU',
         'IND','LAC','LAL','MEM','MIA','MIL','MIN','NOP','NYK','OKC','ORL',
         'PHI','PHX','POR','SAC','SAS','TOR','UTA','WAS']

def getDataSet(dataset):
    df = pd.read_excel(dataset)
    print('Dataset Loaded')
    return df


def SplitTeams(df):
    writer = ExcelWriter(WORKING_DATA_FILE)
    print('Splitting Data')
    for i in teamList:
        df1 = df.loc[df['TEAM_ABBREVIATION'] == i]
        #Format Date and calc rest
        df1['GAME_DATE'] = pd.to_datetime(df1['GAME_DATE_EST'])
        df1['DaysRest'] = df1['GAME_DATE'] - df1['GAME_DATE'].shift(1)
        df1['HomeTeam'] = df1['GAMECODE'].str[12:]
        df1['AwayTeam'] = df1['GAMECODE'].str[9:12]
        df1['HomeIndex'] = df1.apply(getHomeIndex,axis=1)
        df1['AvgPace']= df1['PACE'].expanding().mean()
        df1['AvgORTG'] = df1['OFF_RATING'].expanding().mean()
        df1['AvgDRTG'] = df1['DEF_RATING'].expanding().mean()
        df1['AvgORTG_L5'] = df1['OFF_RATING'].rolling(window=5).mean()
        df1['AvgDRTG_L5'] = df1['DEF_RATING'].rolling(window=5).mean()
        df1['AvgPace']= df1['AvgPace'].shift(1)
        df1['AvgORTG'] = df1['AvgORTG'].shift(1)
        df1['AvgDRTG'] = df1['AvgDRTG'].shift(1)
        df1['AvgORTG_L5'] = df1['AvgORTG_L5'].shift(1)
        df1['AvgDRTG_L5'] = df1['AvgDRTG_L5'].shift(1)

        df1 = df1.loc[df1['HomeTeam'].isin(['ATL','BOS','BKN','CHA','CHI','CLE','DAL','DEN','DET','GSW','HOU',
                 'IND','LAC','LAL','MEM','MIA','MIL','MIN','NOP','NYK','OKC','ORL',
                 'PHI','PHX','POR','SAC','SAS','TOR','UTA','WAS'])]

        df1 = df1.loc[df1['AwayTeam'].isin(['ATL','BOS','BKN','CHA','CHI','CLE','DAL','DEN','DET','GSW','HOU',
                 'IND','LAC','LAL','MEM','MIA','MIL','MIN','NOP','NYK','OKC','ORL',
                 'PHI','PHX','POR','SAC','SAS','TOR','UTA','WAS'])]

        df1.to_excel(writer,i)
    writer.save()
    print('Data Split')


def getFirstSplit(fileName):
    wb = load_workbook(WORKING_DATA_FILE)
    df = pd.read_excel(fileName, sheetname='ATL')
    tabs = wb.get_sheet_names()
    print("Extracting Team Data")
    for j in tabs:
        if j != 'ATL':
            df4 = pd.read_excel(fileName, sheetname=j)
            frames = [df,df4]
            df= pd.concat(frames)
    dfH = df[(df['HomeIndex'] == 0)]
    dfA = df[(df['HomeIndex'] == 1)]
    df1 = pd.merge(dfH, dfA, on='GAMECODE',how='outer')
    df2 = pd.merge(dfA, dfH, on='GAMECODE',how='outer')
    dfList = [df1,df2]
    df3= pd.concat(dfList)
    # print(df3.head())
    df3 = df3.sort_values(by=['GAMECODE','HomeIndex_x'],ascending=[True,True])
    # writer1 = ExcelWriter("Temp_Team_Data.xlsx")
    # df.to_excel(writer1,'Master')
    # writer1.save()
    # writer2 = ExcelWriter("Final_Dataset.xlsx")
    # df3.to_excel(writer2,'Master')
    # writer2.save()
    print("Complete")
    return df3



def getHomeIndex(df):
    if df['HomeTeam'] == df['TEAM_ABBREVIATION']:
        return (0)
    else:
        return (1)

def SplitJointTeams(df):
    writer = ExcelWriter('Split_Teams_2nd_Time.xlsx')
    print('Splitting Data')
    for i in teamList:
        df1 = df.loc[df['TEAM_ABBREVIATION_x'] == i]
        #Format Date and calc rest
        df1['AvgPace_OPP']= df1['AvgPace_y'].expanding().mean()
        df1['AvgORTG_OPP'] = df1['AvgORTG_y'].expanding().mean()
        df1['AvgDRTG_OPP'] = df1['AvgDRTG_y'].expanding().mean()
        df1['AvgORTG_L5_OPP'] = df1['AvgORTG_y'].rolling(window=5).mean()
        df1['AvgDRTG_L5_OPP'] = df1['AvgDRTG_y'].rolling(window=5).mean()
        df1['Opp_ORTG_vs_Avg'] = df1['AvgORTG_OPP'] / df1['LeagueORTG']
        df1['Opp_DRTG_vs_Avg'] = df1['AvgDRTG_OPP'] / df1['LeagueDRTG']
        df1['Opp_Pace_vs_Avg'] = df1['AvgPace_OPP'] / df1['LeaguePace']

        df1.to_excel(writer,i)
    writer.save()
    print('Data Split')

def getSecondSplit(fileName):
    # wb = load_workbook('Split_Teams_2nd_Time.xlsx')
    df = pd.read_excel(fileName, sheetname='ATL')
    tabs = wb.get_sheet_names()
    print("Extracting Team Data")
    for j in tabs:
        if j != 'ATL':
            df1 = pd.read_excel(fileName, sheetname=j)
            frames = [df,df1]
            df= pd.concat(frames)

    df = df.dropna(subset=['AvgORTG_L5_OPP'])
    df = df.sort_values(by=['GAMECODE','HomeIndex_x'],ascending=[True,True])
    df = df[['GAMECODE','TEAM_ABBREVIATION_x','HomeIndex_x','DaysRest_x','AvgPace_x','AvgORTG_x','AvgDRTG_x','AvgORTG_L5_x','AvgDRTG_L5_x'
    ,'TEAM_ABBREVIATION_y','DaysRest_y','AvgPace_y','AvgORTG_y','AvgDRTG_y','AvgORTG_L5_y','AvgDRTG_L5_y','OFF_RATING_x','DEF_RATING_x','PACE_x']]

    writer1 = ExcelWriter("DataForModel.xlsx")
    df.to_excel(writer1,'Master')
    writer1.save()

    print('Data Split')
    return df

def getLeagueAvg(df):
    df['GAMELINK'] = df['GAMECODE'].str[:8]
    df1 = df[['OFF_RATING_x','GAMELINK']]
    df2 = df[['DEF_RATING_x','GAMELINK']]
    df3 = df[['PACE_x','GAMELINK']]
    df1 = df1.groupby(['GAMELINK'],as_index=False)['OFF_RATING_x'].mean()
    df2 = df2.groupby(['GAMELINK'],as_index=False)['DEF_RATING_x'].mean()
    df3 = df3.groupby(['GAMELINK'],as_index=False)['PACE_x'].mean()

    df4 = pd.merge(df1, df2, on='GAMELINK',how='outer')
    df5 = pd.merge(df4, df3, on='GAMELINK',how='outer')
    df5['LeagueORTG'] = df5['OFF_RATING_x'].expanding().mean()
    df5['LeagueDRTG'] = df5['DEF_RATING_x'].expanding().mean()
    df5['LeaguePace'] = df5['PACE_x'].expanding().mean()
    df5['LeagueORTG'] = df5['LeagueORTG'].shift(1)
    df5['LeagueDRTG'] = df5['LeagueDRTG'].shift(1)
    df5['LeaguePace'] = df5['LeaguePace'].shift(1)
    df5 = df5[['GAMELINK','LeagueORTG','LeagueDRTG','LeaguePace']]

    df6 = pd.merge(df, df5, on='GAMELINK',how='outer')

    # print(df6.head())
    writer1 = ExcelWriter("LeagueAvg.xlsx")
    df5.to_excel(writer1,'Master')
    writer1.save()
    return df6
#---------------- Split data out by team and calculate stats ------------------------
# dataset = getDataSet('AllStats_2016.xlsx')
# SplitTeams(dataset)

#---------------- Consume team data from tabs to make one dataset and combine both teams onto one line------------------------
teamdata1 = getFirstSplit(WORKING_DATA_FILE)
teamdata2 = getLeagueAvg(teamdata1)

#---------------- Split data out again to calculate Opponent stats ------------------------
SplitJointTeams(teamdata2)

#---------------- Cosolidate split team data, remove first 6 rows without Last 5 calcs, Filter only needed columns------------------------
# teamdata3 = getSecondSplit('Split_Teams_2nd_Time.xlsx')
