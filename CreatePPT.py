# 2021/09/10 version 1.1
# Author : Vincent

from pptx import Presentation
import datetime
import os

def get_week_of_month():
    today = datetime.datetime.now()
    begin = int(datetime.date(today.year, today.month, 1).strftime("%W"))
    end = int(datetime.date(today.year, today.month, today.day).strftime("%W"))

    return end - begin + 1

def main():
    # set your name here
    name = 'Vincent'
    # you will get the correct weekday, if you create file on Monday(which is the next monday)
    if datetime.datetime.today().weekday() == 1 :
        week = get_week_of_month() - 1
    else:
        week = get_week_of_month()

    file_name = str(datetime.datetime.now().year) + '-' + str(datetime.datetime.now().month) + 'M-' +str(week) + 'W-週報'

    if os.path.isfile(file_name + '.pptx'):
        keep_going = input('The file is already exist, do you want to delete it and create a new file ? (y/n) ')
    if keep_going == 'y':
        # Create pptx object
        prs = Presentation('src/2020-4M-1W-週報.pptx')


        # Set the title of first page 
        title = prs.slides[0].shapes.title
        title.text = file_name
        # Set the subtitle of first page 
        subtitle = prs.slides[0].placeholders[1]
        subtitle.text = name

        prs.save(file_name + '.pptx')
        print('Sate : File Created -> ' + file_name + '.pptx')
    else:
        print("State : Didn't do anything")

if __name__ == '__main__':
    main()