import json
import urllib
import requests
from datetime import datetime, timedelta


def getToday():
    today = datetime.strftime(datetime.now(), '%Y-%m-%d')
    print(today)
    return today

def getYesterday():
   yesterday = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')
   print(yesterday)
   return yesterday

def getTomorrow():
   tomorrow = datetime.strftime(datetime.now() + timedelta(1), '%Y-%m-%d')
   print(tomorrow)
   return tomorrow

def addGames(yesterday, today):
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
    response = urllib.request.urlopen(url)
    data = json.load(response)
    monthsList = {'10':0,'11':1,'12':2, '01':3,'02':4,'03':5,'04':6,'05':7,'06':8}
    gameMonth = yesterday[5:7]
    month = monthsList[gameMonth]
    todaysGames = []
    print(month)
    gameList = (data['lscd'][month]['mscd']['g'])
    for i in range(len(gameList)):
        if gameList[i]['gdte'] == yesterday:
            todaysGames.append(gameList[i]['gid'])
        elif gameList[i]['gdte'] == today:
            break
    for x in todaysGames:
        print (x)
    return todaysGames

def getTodaysGames(yesterday, today):
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
    response = urllib.request.urlopen(url)
    data = json.load(response)
    monthsList = {'10':0,'11':1,'12':2, '01':3,'02':4,'03':5,'04':6,'05':7,'06':8}
    gameMonth = yesterday[5:7]
    month = monthsList[gameMonth]
    todaysGames = []
    gameList = (data['lscd'][month]['mscd']['g'])
    for i in range(len(gameList)):
        if gameList[i]['gdte'] == yesterday:
            game = gameList[i]['gcode']
            game = game[9:17]
            todaysGames.append(game)
        elif gameList[i]['gdte'] == today:
            break
    for x in todaysGames:
        print (x)
    return todaysGames

class getGames(object):
    def __init__(self):
        self.games = []

    def getGamesOCT(self):
        url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
        response = urllib.request.urlopen(url)
        data = json.load(response)

        firstGame = "0011600102"
        gamesOCT = (data['lscd'][0]['mscd']['g'])
        idOCT = []
        for i in range(len(gamesOCT)):
            if gamesOCT[i]['gid'] > firstGame:
                idOCT.append(gamesOCT[i]['gid'])
        count = 0
        for x in idOCT:
            print (x)
            count = count+1
        print (count)
        return idOCT



    def getGamesNOV(self):
        url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
        response = urllib.request.urlopen(url)
        data = json.load(response)

        gamesNOV = (data['lscd'][1]['mscd']['g'])
        idNOV = []
        for i in range(len(gamesNOV)):
                idNOV.append(gamesNOV[i]['gid'])
        count = 0
        for x in idNOV:
            print (x)
            count = count+1
        print (count)
        return idNOV


    def getGamesDEC(self):
        url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
        response = urllib.request.urlopen(url)
        data = json.load(response)

        gamesDEC = (data['lscd'][2]['mscd']['g'])
        idDEC = []
        for i in range(len(gamesDEC)):
                idDEC.append(gamesDEC[i]['gid'])
        count = 0
        for x in idDEC:
            print (x)
            count = count+1
        print (count)
        return idDEC

    def getGamesJAN(self):
        url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
        response = urllib.request.urlopen(url)
        data = json.load(response)

        today = "2017-01-04"
        gamesJAN = (data['lscd'][3]['mscd']['g'])
        idJAN = []
        for i in range(len(gamesJAN)):
            if gamesJAN[i]['gdte'] != today:
                idJAN.append(gamesJAN[i]['gid'])
            else:
                break
        count = 0
        for x in idJAN:
            print (x)
            count = count+1
        print (count)
        return idJAN



