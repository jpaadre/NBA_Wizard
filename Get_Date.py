from datetime import datetime, timedelta


def getToday():
    today = datetime.strftime(datetime.now(), '%Y-%m-%d')
    return today
    print(today)

def getYesterday():
   yesterday = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')
   return yesterday
   print(yesterday)

