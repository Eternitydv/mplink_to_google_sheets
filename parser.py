import requests
import json
from lxml import html
import pickle
from tkinter import *
from tkinter import ttk
import os
import numpy as np
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import httplib2 
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import statistics
import pygsheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CREDENTIALS_FILE = 'D:\\Useful misc\\мплинкпарсер\\mplink_to_google_sheets\\match-results-19f739c32961.json'

class ui:
    link = 'https://osu.ppy.sh/community/matches/71542504'
    #spreadsheetId = '1Kst83QCbmRISmUt0zDO7938kZ9SdUczlpGlVh_qIVRU'
    spreadsheetId = '11fn6U7RhMPTD69EgZTgwvU_JSEOZt8QINd9mJ8Q8Y9s'
    sheetname = 'Week 2(1)'
    startColumnIndex = 7
    endColumnIndex = 67
    startRowIndex = 2
    endRowIndex = 18
    names = []
    diff = {'Easy' : 600000, 'Medium' : 500000, 'Hard' : 400000, 'Markrum' : 150000}
    def __init__(self, root):
        root.title("Булщит в шит")
        root.geometry("500x100")
        mainframe = ttk.Frame(root)
        mainframe.pack(fill=BOTH, expand=True)

        frame_top = Frame(mainframe)
        frame_top.pack()
        ttk.Label(frame_top, text="Айди дока: ").pack(side=LEFT)
        sheet_id = ttk.Entry(frame_top, width=50)
        sheet_id.pack(side=LEFT)
        sheet_id.insert(0, self.spreadsheetId)

        frame_mid = Frame(mainframe)
        frame_mid.pack()
        ttk.Label(frame_mid, text = "Название листа: Week ").pack(side=LEFT)
        sheetname_entry1 = ttk.Entry(frame_mid, width = 2, textvariable = self.sheetname)
        sheetname_entry1.pack(side=LEFT)
        ttk.Label(frame_mid, text = "(").pack(side=LEFT)
        sheetname_entry2 = ttk.Entry(frame_mid, width = 2, textvariable = self.sheetname)
        sheetname_entry2.pack(side=LEFT)
        ttk.Label(frame_mid, text = ")").pack(side=LEFT)
        
        frame_third = Frame(mainframe)
        frame_third.pack()
        ttk.Label(frame_third, text="MP link: ").pack(side=LEFT)
        link_space = ttk.Entry(frame_third, width=50, textvariable = self.link)
        link_space.pack(side=LEFT)

        frame_bot = Frame(mainframe)
        frame_bot.pack()
        ttk.Button(frame_bot, text="Отправить результаты", command = lambda: self.get_shit_to_sheet(link_space, sheetname_entry1, sheetname_entry2, sheet_id)).pack(side=LEFT)
        self.status = ttk.Label(frame_bot, text='Ready')
        self.status.pack(side=LEFT)

        root.bind("<Return>", lambda: self.get_shit_to_sheet(link_space, sheetname_entry1, sheetname_entry2, sheet_id))
        link_space.focus()

    def get_username_from_id(self, user_id):
        url = "https://osu.ppy.sh/users/" + str(user_id)
        page = requests.get(url)
        tree = html.fromstring(page.content)
        results = tree.xpath("//script[@id='json-user']/text()")
        data = json.loads(str(results[0]))
        return data["username"]

    def parse_the_link(self, link):
        numberOfPeople = 0
        print('started parsing')
        self.status.config(text = 'Parsing the mp link')
        page = requests.get(link)
        tree = html.fromstring(page.content)
        results = tree.xpath("//script[@id='json-events']/text()")
        data = json.loads(str(results[0]))
        dict_of_scores = {}
        counter = 0
        mappool_size = 0
        for item in data["events"]:
            if "game" in item.keys():
                #print(json.dumps(item,indent=4))
                for score in item["game"]["scores"]:
                    name = score["user_id"]
                    if name not in dict_of_scores.keys():
                        numberOfPeople+=1
                        dict_of_scores[name] = [[], [], []]
                        for i in range(counter):
                            dict_of_scores[name][0].append(0)
                            dict_of_scores[name][1].append(0) 
                            dict_of_scores[name][2].append(0) 
                            
                    dict_of_scores[name][0].append(score["score"])
                    dict_of_scores[name][1].append('%.2f'%(score["accuracy"]*100))

                    if 'DT' in score['mods']:
                        dict_of_scores[name][2].append('DT')
                    elif 'HR' in score['mods']:
                        dict_of_scores[name][2].append('HR')
                    elif 'HD' in score['mods']:
                        dict_of_scores[name][2].append('HD')
                    else:
                        dict_of_scores[name][2].append('NM')
                counter+=1

        temp = []
        
        for score in dict_of_scores.values():
            temp.append(len(score[0]))

        self.mappool_size = int(statistics.median(temp))
        print('mappool size = ' + str(self.mappool_size))

        self.endColumnIndex = self.startColumnIndex + numberOfPeople*2 - 1
        print(self.endColumnIndex)
        counter = 0

        for scores in dict_of_scores.values(): #will readjust if someone had played the maps after the end of the lobby
            if len(scores[0]) < self.mappool_size:
                for i in range(len(scores[0]), self.mappool_size):
                    scores[0].append(0)
                    scores[1].append(0)
                    scores[2].append(0)

            if len(scores[0]) > self.mappool_size:
                for i in range(self.mappool_size):
                    if scores[0][i] == 0:
                        scores[0][i] = scores[0].pop(self.mappool_size)
                        scores[1][i] = scores[1].pop(self.mappool_size)
                        scores[2][i] = scores[2].pop(self.mappool_size)
                        

        self.endRowIndex = self.startRowIndex + self.mappool_size + 1
        print('endRowIndex = ' + str(self.endRowIndex))

        for id in list(dict_of_scores.keys()):
            self.names.append(self.get_username_from_id(id))

        print('done with parsing')
        self.status.config(text = 'Finished parsing the mp link')
        return dict_of_scores

    def get_shit_to_sheet(self, link_space, sheetname_entry1, sheetname_entry2, sheet_id):
        self.link = link_space.get()
        self.sheetname = 'Week {}({})'.format(sheetname_entry1.get(), sheetname_entry2.get())
        if sheet_id.get() != None:
            self.spreadsheetId = sheet_id.get()

        results = self.parse_the_link(self.link)

        self.status.config(text = 'Authorizing gsheets api') 
        gc = pygsheets.authorize(service_file=CREDENTIALS_FILE)
        sprdsheet = gc.open_by_key(self.spreadsheetId)
        sheet = sprdsheet.worksheet_by_title(self.sheetname)

        print(self.spreadsheetId)
        print(self.sheetname)

        difficulty = sheet.cell('C23').value_unformatted
        print(difficulty)
        self.status.config(text = 'Prepping data for the api calls')
        data_dump = []
        color_ref = []
        i = 0
        #prepping data dumps for api calls
        for scores in results.values():
            scores[0].append('%.2f'%(sum(scores[0])/self.mappool_size/self.diff[difficulty]))
            scores[0].insert(0, self.names[i])
            data_dump.append(scores[0])
            scores[1].insert(0, None)
            data_dump.append(scores[1])
            color_ref.append(scores[2])
            color_ref.append(scores[2])
            i+=1

        sheet.update_value((self.endRowIndex, 6), 'Match cost: ')

        values = list(results.values())
        
        data_range = pygsheets.datarange.DataRange(start=(self.startRowIndex, self.startColumnIndex), end=(self.endRowIndex, self.endColumnIndex), worksheet=sheet)
        color_range = pygsheets.datarange.DataRange(start=(self.startRowIndex + 1, self.startColumnIndex), end=(self.endRowIndex - 1, self.endColumnIndex), worksheet=sheet)
        
        le_epic_color_range = pygsheets.datarange.DataRange(start = (self.startRowIndex+50, self.startColumnIndex), end = (self.endRowIndex+50, self.endColumnIndex), worksheet=sheet)
        self.status.config(text = 'Inputting score data')

        sheet.update_values(crange = data_range.range, values = data_dump, majordim = 'COLUMNS', parse=False)

        sheet.update_values(crange = le_epic_color_range.range, values = color_ref, majordim = 'COLUMNS', parse=False)

        self.status.config(text = 'Epic color programming')
        
        #adding rules for coloring cells
        sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-1, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
         format={'backgroundColor' : {'red' : 234/255, 'green' : 153/255, 'blue' : 153/255, 'alpha' : 1}}, condition_values=['=(G{}=\"HR\")'.format(self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-1, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
         format={'backgroundColor' : {'red' : 180/255, 'green' : 167/255, 'blue' : 214/255, 'alpha' : 1}}, condition_values=['=(G{}=\"DT\")'.format(self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-1, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
         format={'backgroundColor' : {'red' : 255/255, 'green' : 229/255, 'blue' : 153/255, 'alpha' : 1}}, condition_values=['=(G{}=\"HD\")'.format(self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-1, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
         format={'backgroundColor' : {'red' : 250/255, 'green' : 238/255, 'blue' : 243/255, 'alpha' : 1}}, condition_values=['=(G{}=\"NM\")'.format(self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-1, self.endColumnIndex), condition_type='CUSTOM_FORMULA', 
          format={'backgroundColor' : {'red' : 75/255, 'green' : 75/255, 'blue' : 75/255, 'alpha' : 1}}, condition_values=['=(G{}=0)'.format(self.startRowIndex+50)])

        self.status.config(text = 'Finished updating {}'.format(self.sheetname))
        self.names.clear()
        print('finished updating {}'.format(self.sheetname))

root = Tk()
ui(root)
root.mainloop()