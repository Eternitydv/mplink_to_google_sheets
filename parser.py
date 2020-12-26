from bs4 import BeautifulSoup
import requests
import json
from lxml import html
import pickle
from tkinter import *
from tkinter import ttk
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import httplib2 
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import statistics
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

class ui:
    link = 'https://osu.ppy.sh/community/matches/72098518'
    #spreadsheetId = '1Kst83QCbmRISmUt0zDO7938kZ9SdUczlpGlVh_qIVRU'
    spreadsheetId = '11fn6U7RhMPTD69EgZTgwvU_JSEOZt8QINd9mJ8Q8Y9s'
    sheetname = 'Week 3(2)'
    startColumnIndex = 7
    endColumnIndex = 67
    startRowIndex = 3
    endRowIndex = 18
    numberOfPeople = 0
    incr = {}
    diff = {'Easy' : 600000, 'Medium' : 500000, 'Hard' : 400000, 'Markrum' : 150000}

    def __init__(self, root):
        root.title("Булщит в шит")
        root.geometry("500x100")
        mainframe = ttk.Frame(root, padding = "3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        sheet_id = ttk.Entry(mainframe, width = 50, textvariable = self.spreadsheetId)
        sheet_id.grid(column = 1, row = 1)
        sheet_id.insert(0, self.spreadsheetId)
        ttk.Label(mainframe, text="Айди дока: ").grid(column = 0, row = 1)

        ttk.Label(mainframe, text = "Название листа: ").grid(column=0, row=3)
        sheetname_entry = ttk.Entry(mainframe, width = 20, textvariable = self.sheetname)
        sheetname_entry.grid(column = 1, row = 3, sticky=(W, E))
        sheetname_entry.insert(0, 'Week ()')

        link_space = ttk.Entry(mainframe, width = 50, textvariable = self.link)
        link_space.grid(column = 1, row = 5, sticky=(W, E))
        ttk.Label(mainframe, text="MP link: ").grid(column=0, row = 5)

        ttk.Button(mainframe, text="Отправить результаты", command = lambda: self.get_shit_to_sheet(link_space, sheetname_entry, sheet_id)).grid(column = 0, row = 6, sticky =W)
        root.bind("<Return>", lambda: self.get_shit_to_sheet(link_space, sheetname_entry, sheet_id))
        link_space.focus()

    def get_username_from_id(self, user_id):
        url = "https://osu.ppy.sh/users/" + str(user_id)
        page = requests.get(url)
        tree = html.fromstring(page.content)
        results = tree.xpath("//script[@id='json-user']/text()")
        data = json.loads(str(results[0]))
        return data["username"]

    def parse_the_link(self, link):
        print('started parsing')
        global numberOfPeople, endColumnIndex, endRowIndex
        page = requests.get(link)
        tree = html.fromstring(page.content)
        results = tree.xpath("//script[@id='json-events']/text()")
        data = json.loads(str(results[0]))
        lobby_dict = {}
        dict_of_scores = {}
        counter = 0
        mappool_size = 0
        for item in data["events"]:
            if "game" in item.keys():
                #print(json.dumps(item,indent=4))
                for score in item["game"]["scores"]:
                    name = score["user_id"]
                    if name not in dict_of_scores.keys():
                        dict_of_scores[name] = [[], []]
                        for i in range(counter):
                            dict_of_scores[name][0].append(0)
                            dict_of_scores[name][1].append(0)
                    dict_of_scores[name][0].append(score["score"])
                    dict_of_scores[name][1].append('%.2f'%(score["accuracy"]*100))
                counter+=1
        temp = []
        for score in dict_of_scores.values():
            temp.append(len(score[0]))
        self.mappool_size = int(statistics.median(temp))
        print('mappool size = ' + str(self.mappool_size))

        list_of_ids = list(dict_of_scores.keys())
        names = []
        self.numberOfPeople = len(list_of_ids)

        self.endColumnIndex = self.startColumnIndex + self.numberOfPeople*2
        counter = 0
        for scores in dict_of_scores.values():
            print(scores[0])
            if len(scores[0]) < self.mappool_size:
                for i in range(len(scores[0]), self.mappool_size):
                    scores[0].append(0)
                    scores[1].append(0)
            if len(scores[0]) > self.mappool_size:
                for i in range(self.mappool_size):
                    if scores[0][i] == 0:
                        scores[0][i] = scores[0].pop(self.mappool_size)
                        scores[1][i] = scores[1].pop(self.mappool_size)

        self.endRowIndex = self.startRowIndex + self.mappool_size +  1

        for id in list_of_ids:
            names.append(self.get_username_from_id(id))
        
        for i in range(len(list_of_ids)):
            dict_of_scores[names[i]] = dict_of_scores.pop(list_of_ids[i])
        print('done with parsing')
        return dict_of_scores

    def fill_the_incr(self):
        for i in range(71, 71 + self.numberOfPeople*2 +1):
            if i < 89:
                self.incr[chr(i)] = chr(i + 2)
            else:
                if i < 91:
                    self.incr[chr(i)] = "A" + chr(i-64)
                else:
                    ch = "A" + chr(i-66)
                    self.incr[ch] = "A" + chr(i-64)
        print(self.incr)

    def get_shit_to_sheet(self, link_space, sheetname_entry, sheet_id):
        self.link = link_space.get()
        self.sheetname = sheetname_entry.get()
        if sheet_id.get() != None:
            self.spreadsheetId = sheet_id.get()
        results = self.parse_the_link(self.link)
        self.fill_the_incr()
        CREDENTIALS_FILE = 'D:/Useful misc/мплинкпарсер/mplink_to_google_sheets/match-results-19f739c32961.json'
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
        ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
        service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

        service = build('sheets', 'v4', credentials=credentials)
        spreadsheet = service.spreadsheets().get(spreadsheetId = self.spreadsheetId).execute()
        difficulty = service.spreadsheets().values().get(spreadsheetId = self.spreadsheetId, 
        range = self.sheetname + "!C23", valueRenderOption='UNFORMATTED_VALUE').execute()['values'][0][0]
        print(difficulty)
        divider = self.diff[difficulty]
        for scores in results.values():
            scores[0].append('%.2f'%(sum(scores[0])/self.mappool_size/divider))
            print(scores[0][-1])
        print("end row index " + str(self.endRowIndex))
        self.columnToWrite = "G"
        print('starting to fill the sheet')
        for k in results.items():
            batch_update_spreadsheet_request_body = {
                "valueInputOption": "USER_ENTERED",
                "data" : 
                [
                    {
                        "range" : self.sheetname + "!" + self.columnToWrite +"3:AL25",
                        "majorDimension" : "COLUMNS",
                        "values" : 
                        [
                            k[1][0], k[1][1]
                        ]
                    },
                    {
                        "range" : self.sheetname + "!"+ self.columnToWrite +"2",
                        "majorDimension" : "rows",
                        "values" : 
                        [
                            [k[0]]
                        ]
                    },
                    {
                        "range" : self.sheetname + "!F" + str(self.endRowIndex-1),
                        "majorDimension" : "rows",
                        "values" : 
                        [
                            ['Match cost: ']
                        ]
                    }
                ]
            }
            res = service.spreadsheets().values().batchUpdate(spreadsheetId = self.spreadsheetId,
            body = batch_update_spreadsheet_request_body).execute()
            self.columnToWrite = self.incr[self.columnToWrite]
            print('done with ' + k[0])

root = Tk()
ui(root)
root.mainloop()