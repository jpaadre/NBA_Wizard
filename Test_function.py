import Schedule_Reader
import Stat_Finder
import Get_Summary_Stats
import getOpponentAvg
import Get_League_Avg

#Get dates needed to get games
yesterday = Schedule_Reader.getYesterday()
today = Schedule_Reader.getToday()
tomorrow = Schedule_Reader.getTomorrow()
'''
#Get yesterdays games to be added to spreadsheet
games = Schedule_Reader.addGames(yesterday, today)

#Get stats for yesterdays games and add them to team spreadsheets
Stat_Finder.getStats(games)

#Calculate the rest of yesterday's games
Stat_Finder.calcRest()

#Calculate the rest detail of yesterday's games
Stat_Finder.getRestDetail()

'''
#Update the summary stats tabs with updated summary data
Get_Summary_Stats.calcORTG()
Get_Summary_Stats.calcDRTG()
Get_Summary_Stats.calcPace()
Get_League_Avg.getLeagueAvg()

#Create a list of games being played today
todaysGames = Schedule_Reader.getTodaysGames(today,tomorrow)
#Pull stats from summary tabs for the teams playing today
ORTGs = getOpponentAvg.getORTG(todaysGames)
proj = getOpponentAvg.ProjectScores(todaysGames,ORTGs)

#GetProjection.getTodaysStats(todaysGames)

#Uses summary stats to project games

#Scrapes betting lines from online and puts them in Excel

#Compares projected score verus betting lines

#Gives each projection a grade based on probability of winning(Monte Carlo)?
#and deviation from betting line

#Check to see how projection did and track win/losses in Excel per game

#Develop new models

