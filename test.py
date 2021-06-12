# do not delete

from datetime import date
import pandas as pd


def get_yesterday_date(search):                                   #returns yesterday date

    if search.group() == 'yesterday':
        today = date.today()
        

        date_red = int(today.strftime('%d'))
        date_red = date_red - 1 

        return today.strftime(f'{date_red}-%m-%y')


def status_attdendance2(words):                                    #returns totall number of present and absent students
       
        total_present = 0
        total_absent = 0

        for row in words:
            if row == "P":
                total_present += 1
            else:
                total_absent += 1

        return total_present, total_absent










