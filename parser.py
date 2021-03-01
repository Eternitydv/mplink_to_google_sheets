import requests
import json
from lxml import html
from tkinter import *
from tkinter import ttk
import os
import statistics
import pygsheets
import pickle
import os.path
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
#CREDENTIALS_FILE = os.getcwd()+'\match-results-19f739c32961.json'
CREDENTIALS_FILE = os.getcwd()+'\match-results-19f739c32961.json'

class ui:
    link = 'https://osu.ppy.sh/community/matches/71707670'
    #spreadsheetId = '1D0Wa7xBdwjnwmIz5ACsNN775s_zy3a4EhftO39Vl_U4'
    spreadsheetId = '1Kst83QCbmRISmUt0zDO7938kZ9SdUczlpGlVh_qIVRU'
    sheetname = 'Week 2(2)'
    startColumnIndex = 7
    endColumnIndex = 67
    startRowIndex = 2
    endRowIndex = 18
    names = []
    k = 0
    diff = {'Easy' : 600000, 'Medium' : 500000, 'Qualis' : 450000, 'Hard' : 400000, 'Insane' : 300000, 'Markrum' : 150000}
    def __init__(self, root):
        k = IntVar()
        root.title("Булщит в шит")
        root.geometry("450x100")
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
        sheetname_entry1 = ttk.Entry(frame_mid, width = 2)
        sheetname_entry1.pack(side=LEFT)
        ttk.Label(frame_mid, text = "(").pack(side=LEFT)
        sheetname_entry2 = ttk.Entry(frame_mid, width = 2)
        sheetname_entry2.pack(side=LEFT)
        ttk.Label(frame_mid, text = ")").pack(side=LEFT)
        w = Checkbutton(frame_mid, text = 'Квалы', variable = k, onvalue = 1, offvalue = 0)
        w.pack(side = LEFT)
        
        frame_third = Frame(mainframe)
        frame_third.pack()
        ttk.Label(frame_third, text="MP link: ").pack(side=LEFT)
        link_space = ttk.Entry(frame_third, width=50, textvariable = self.link)
        link_space.pack(side=LEFT)

        frame_bot = Frame(mainframe)
        frame_bot.pack()
        ttk.Button(frame_bot, text="Отправить результаты", command = lambda: self.to_sheet(link_space, sheetname_entry1, sheetname_entry2, sheet_id, k, add = False)).pack(side=LEFT)
        ttk.Button(frame_bot, text="Добавить результаты", command = lambda: self.to_sheet(link_space, sheetname_entry1, sheetname_entry2, sheet_id, k, add = True)).pack(side=LEFT)
        self.status = ttk.Label(frame_bot, text='Ready')
        self.status.pack(side=LEFT)

        link_space.focus()

    def get_username_from_id(self, user_id):
        url = "https://osu.ppy.sh/users/" + str(user_id)
        page = requests.get(url)
        tree = html.fromstring(page.content)
        results = tree.xpath("//script[@id='json-user']/text()")
        data = json.loads(str(results[0]))
        return data["username"]

    def fix_score_list(self, dict_of_scores):
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
        return dict_of_scores

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
        l = 0
        beatmap_ids = []
        for item in data["events"]:
            if item['detail']['type'] == 'player-joined':
                if item['user_id'] in dict_of_scores.keys():
                    for r in range(counter - len(dict_of_scores[item['user_id']][0])):
                        dict_of_scores[item['user_id']][0].append(0)
                        dict_of_scores[item['user_id']][1].append(0) 
                        dict_of_scores[item['user_id']][2].append(0) 

            #print(json.dumps(item,indent=4))
            if "game" in item.keys() and item['game']['scores']:

                if item['game']['beatmap']['id'] not in beatmap_ids:
                    beatmap_ids.append(item['game']['beatmap']['id'])
    
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

                elif item['game']['beatmap']['id'] == beatmap_ids[-1]:
                    for score in item["game"]["scores"]:
                        name = score["user_id"]
                        if name not in dict_of_scores.keys():
                            numberOfPeople+=1
                            dict_of_scores[name] = [[], [], []]
                            for i in range(counter):
                                dict_of_scores[name][0].append(0)
                                dict_of_scores[name][1].append(0) 
                                dict_of_scores[name][2].append(0) 
                                
                        dict_of_scores[name][0][-1] = score["score"]
                        dict_of_scores[name][1][-1] = '%.2f'%(score["accuracy"]*100)

                        if 'DT' in score['mods']:
                            dict_of_scores[name][2][-1] = 'DT'
                        elif 'HR' in score['mods']:
                            dict_of_scores[name][2][-1] = 'HR'
                        elif 'HD' in score['mods']:
                            dict_of_scores[name][2][-1] = 'HD'
                        else:
                            dict_of_scores[name][2][-1] = 'NM'
                
        temp = []
        print(dict_of_scores)

        for score in dict_of_scores.values():
            temp.append(len(score[0]))
        self.mappool_size = counter
        print('mappool size = ' + str(self.mappool_size))

        self.endColumnIndex = self.startColumnIndex + numberOfPeople*2 - 1
        print(self.endColumnIndex)
        
        dict_of_scores = self.fix_score_list(dict_of_scores)         

        self.endRowIndex = self.startRowIndex + self.mappool_size + 2

        print('endRowIndex = ' + str(self.endRowIndex))

        for id in list(dict_of_scores.keys()):
            self.names.append(self.get_username_from_id(id))

        print('done with parsing')
        self.status.config(text = 'Finished parsing the mp link')
        return dict_of_scores

    def add_mod_color_rules(self, sheet, _range):
        #adding rules for coloring cells
        if self.startColumnIndex + 64 < 91:
            column = chr(self.startColumnIndex + 64)
        elif self.startColumnIndex + 64 < 117:
            column = "A" + chr(self.startColumnIndex + 38)
        elif self.startColumnIndex + 64 < 143:
            column = 'B' + chr(self.startColumnIndex + 12)
        elif self.startColumnIndex + 64 < 169:
            column = 'C' + chr(self.startColumnIndex - 14)
        #============================NEW FORMAT==============================================================================================================
        sheet.add_conditional_formatting(start = _range.start_addr, end = _range.end_addr, condition_type='CUSTOM_FORMULA',
            format={'textFormat' : {'foregroundColor' : {'red' : 224/255, 'green' : 102/255, 'blue' : 102/255, 'alpha' : 1}}}, condition_values=['=({}{}=\"HR\")'.format(column, self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = _range.start_addr, end = _range.end_addr, condition_type='CUSTOM_FORMULA',
            format={'textFormat' : {'foregroundColor' : {'red' : 142/255, 'green' : 124/255, 'blue' : 195/255, 'alpha' : 1}}}, condition_values=['=({}{}=\"DT\")'.format(column, self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = _range.start_addr, end = _range.end_addr, condition_type='CUSTOM_FORMULA',
            format={'textFormat' : {'foregroundColor' : {'red' : 255/255, 'green' : 217/255, 'blue' : 102/255, 'alpha' : 1}}}, condition_values=['=({}{}=\"HD\")'.format(column, self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = _range.start_addr, end = _range.end_addr, condition_type='CUSTOM_FORMULA',
            format={'textFormat' : {'foregroundColor' : {'red' : 109/255, 'green' : 158/255, 'blue' : 235/255, 'alpha' : 1}}}, condition_values=['=({}{}=\"NM\")'.format(column, self.startRowIndex+50)])

        sheet.add_conditional_formatting(start = _range.start_addr, end = _range.end_addr, condition_type='CUSTOM_FORMULA', 
            format={'textFormat' : {'foregroundColor' : {'red' : 75/255, 'green' : 75/255, 'blue' : 75/255, 'alpha' : 1}}}, condition_values=['=({}{}=0)'.format(column, self.startRowIndex+50)])

        # sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-2, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
        #  format={'backgroundColor' : {'red' : 234/255, 'green' : 153/255, 'blue' : 153/255, 'alpha' : 1}}, condition_values=['=(G{}=\"HR\")'.format(self.startRowIndex+50)])

        # sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-2, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
        #  format={'backgroundColor' : {'red' : 180/255, 'green' : 167/255, 'blue' : 214/255, 'alpha' : 1}}, condition_values=['=(G{}=\"DT\")'.format(self.startRowIndex+50)])

        # sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-2, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
        #  format={'backgroundColor' : {'red' : 255/255, 'green' : 229/255, 'blue' : 153/255, 'alpha' : 1}}, condition_values=['=(G{}=\"HD\")'.format(self.startRowIndex+50)])

        # sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-2, self.endColumnIndex), condition_type='CUSTOM_FORMULA',
        #  format={'backgroundColor' : {'red' : 250/255, 'green' : 238/255, 'blue' : 243/255, 'alpha' : 1}}, condition_values=['=(G{}=\"NM\")'.format(self.startRowIndex+50)])

        # sheet.add_conditional_formatting(start = (self.startRowIndex + 1, self.startColumnIndex), end = (self.endRowIndex-2, self.endColumnIndex), condition_type='CUSTOM_FORMULA', 
        #   format={'backgroundColor' : {'red' : 75/255, 'green' : 75/255, 'blue' : 75/255, 'alpha' : 1}}, condition_values=['=(G{}=0)'.format(self.startRowIndex+50)])

    def update_range_data(self, sheet):
        cell_list = sheet.range(crange='3:3', returnas='matrix')[0]
        del cell_list[0:6]
        cell_list = [i for i in cell_list if i != '']
        self.startColumnIndex += len(cell_list)
        self.endColumnIndex = self.startColumnIndex + 1
        print("self.startColumnIndex = " + str(self.startColumnIndex))

        cell_list = sheet.range(crange='E3:E30', returnas='matrix')
       # print(cell_list)
        cell_list = [i for i in cell_list if i != ['']]
       # print(cell_list)
        self.endRowIndex = self.startRowIndex + len(cell_list) + 1
        print(self.endRowIndex)

    def to_sheet(self, link_space, sheetname_entry1, sheetname_entry2, sheet_id, w, add):
        print(w.get())
        self.link = link_space.get()
        if w.get() == 1:
            self.sheetname = 'Квалы'
        else: 
            self.sheetname = 'Week {}({})'.format(sheetname_entry1.get(), sheetname_entry2.get())

        print(self.sheetname)
        if sheet_id.get() != None:
            self.spreadsheetId = sheet_id.get()

        results = self.parse_the_link(self.link)

        self.status.config(text = 'Authorizing gsheets api') 

        gc = pygsheets.authorize(service_file=CREDENTIALS_FILE)
        spreadsheet = gc.open_by_key(self.spreadsheetId)
        sheet = spreadsheet.worksheet_by_title(self.sheetname)

        if add:
            self.update_range_data(sheet)
            results = self.fix_score_list(results)
        print(self.startColumnIndex)

        print(self.spreadsheetId)
        print(self.sheetname)
        print("endRowIndex = " + str(self.endRowIndex))
        print('startRowIndex = ' + str(self.startRowIndex))

        #=================NEW FORMAT==============================
        difficulty = sheet.cell((self.endRowIndex, 4)).value_unformatted
        #=================OLD FORMAT================================
        #difficulty = sheet.cell((self.endRowIndex, 3)).value_unformatted
        print(difficulty)
        self.status.config(text = 'Prepping data for the api calls')
        data_dump = []
        color_ref = []
        i = 0
        print(results)
        #prepping data dumps for api calls
        for scores in results.values():
            match_cost = '%.2f'%(sum(scores[0])/self.mappool_size/self.diff[difficulty])
            scores[0].append(None)
            scores[0].append(match_cost)
            scores[0].insert(0, self.names[i])
            data_dump.append(scores[0])
            scores[1].insert(0, None)
            scores[1].append(None)
            data_dump.append(scores[1])
            color_ref.append(scores[2])
            color_ref.append(scores[2])
            i+=1
        self.status.config(text = 'Api calls')

        sheet.update_value((self.endRowIndex, 6), 'Match cost: ')

        values = list(results.values())
        
        data_range = pygsheets.datarange.DataRange(start=(self.startRowIndex, self.startColumnIndex), end=(self.endRowIndex, self.endColumnIndex), worksheet=sheet)
        color_range = pygsheets.datarange.DataRange(start=(self.startRowIndex + 1, self.startColumnIndex), end=(self.endRowIndex - 2, self.endColumnIndex), worksheet=sheet)

        print('data_dump: ' + str(data_dump))

        le_epic_color_range = pygsheets.datarange.DataRange(start = (self.startRowIndex+50, self.startColumnIndex), end = (self.endRowIndex+50, self.endColumnIndex), worksheet=sheet)
        self.status.config(text = 'Inputting score data')

        sheet.update_values(crange = data_range.range, values = data_dump, majordim = 'COLUMNS', parse=False)

        sheet.update_values(crange = le_epic_color_range.range, values = color_ref, majordim = 'COLUMNS', parse=False)

        self.status.config(text = 'Epic color programming')
        
        self.add_mod_color_rules(sheet, color_range)

        print('finished updating {}'.format(self.sheetname))
        self.status.config(text = 'Updating stats')
        if add:
            self.update_stats_add(spreadsheet, data_dump)
        else:
            self.update_stats_initial(spreadsheet, data_dump)
        self.status.config(text = 'Finished updating {}'.format(self.sheetname))
        self.names.clear()
        self.startColumnIndex = 7
        self.startRowIndex = 2

    def update_stats_add(self, spreadsheet, data_dump):
        stats = spreadsheet.worksheet_by_title('Stats')
        cell_list = stats.range(crange='1:1', returnas='matrix')[0]
        weeks = [i for i in cell_list if i]

        cell_list = stats.range(crange='A1:A50', returnas='matrix')
        names_in_stats = [i[0] for i in cell_list if i]
        print(names_in_stats)
        col = week.index([self.sheetname]) + 1
        if data_dump[0][0] in names_in_stats:
            row = names_in_stats.index(data_dump[0][0]) + 2
        else:
            row = len(names_in_stats) + 2
            stats.update_value(addr=(row, 1), val=data_dump[0][0])
        stats.update_value(addr=(row, column), val = data_dump[0][-1])


    def update_stats_initial(self, spreadsheet, data_dump):
        stats = spreadsheet.worksheet_by_title('Stats')
        cell_list = stats.range(crange='1:1', returnas='matrix')[0]
        cell_list = [i for i in cell_list if i]
        if self.sheetname in cell_list:
            column = cell_list.index(self.sheetname) + 1
        else:
            column = len(cell_list) + 1


        cell_list = stats.range(crange='A1:A50', returnas='matrix')
        del cell_list[0]
        names_in_stats = [i[0] for i in cell_list if i]
        results = {}
        for i in data_dump[::2]:
            results[i[0]] = i[-1]
        data = [self.sheetname]
        name_data = []
        for i in list(results.keys()):
            if i not in names_in_stats:
                name_data.append(i)
                names_in_stats.append(i)
        for i in names_in_stats:
            if i in list(results.keys()):
                data.append(results[i])
            else:
                data.append('')

        data_range = pygsheets.datarange.DataRange(start=(1, column), end=(len(data) + 1, column), worksheet=stats)
        name_range = pygsheets.datarange.DataRange(start=(len(names_in_stats)+2, 1), end=(len(names_in_stats) + 2 + len(name_data), 1), worksheet=stats)
        stats.update_values(crange = data_range.range, values = [data], majordim = 'COLUMNS', parse=True)
        stats.update_values(crange = name_range.range, values = [name_data], majordim = 'COLUMNS', parse=False)
        

    # def update_stats(self, spreadsheet, results):
    #     sheet = spreadsheet.worksheet_by_title("Stats")
    #     name_range = sheet.range(crange="1:1", returnas="matrix")[0]
    #     name_range = [i for i in name_range if i !='']
    #     print(name_range)
    #     data = []
    #     names_to_submit = list(results.keys())
    #     for name in name_range:
    #         if name in names_to_submit:
    #             data.append(results[name][0][-1])
    #         else:
    #             data.append(None)
    #     row = sheet.
    #     submit_range = pygsheets.datarange.DataRange(start = (, 2), end = (, 2 + len(name_range)), worksheet=sheet)
    #     sheet.update_values(crange = )

# def backport_stats():
#     gc = pygsheets.authorize(service_file=CREDENTIALS_FILE)
#     spreadsheet = gc.open_by_key('1Kst83QCbmRISmUt0zDO7938kZ9SdUczlpGlVh_qIVRU')
#     weeks = []
#     weeks.append(spreadsheet.worksheet_by_title('Week 1(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 2(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 2(2)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 3(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 3(2)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 4(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 5(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 6(1)'))
#     weeks.append(spreadsheet.worksheet_by_title('Week 6(2)'))
#     results = {}
#     counter = 0
#     week_titles = []
#     for i in weeks:
#         week_titles.append(i.title)
#         print('working on ' + i.title)
#         temp = get_stats(i)
#         print('temp: ' + str(temp))
#         for key in list(results.keys()):
#             if key not in list(temp.keys()):
#                 results[key].append('')
#         for k, v in temp.items():
#             if k not in results:
#                 results[k] = []
#                 for j in range(counter):
#                     results[k].append('')
#                 results[k].append(v)
#             else:
#                 print(k)
#                 print(v)
#                 results[k].append(v)
#         print('results: ' + str(results))
#         counter+=1
        
#     stats = spreadsheet.worksheet_by_title('Stats')
#     print(len(list(results.keys())))
#     number_of_people = len(list(results.keys()))
#     name_range = pygsheets.datarange.DataRange(start=(1, 2), end=(1, 2 + number_of_people), worksheet=stats)
#     data_range = pygsheets.datarange.DataRange(start=(3, 1), end=(11, 2 + number_of_people), worksheet=stats)
#     name_data = []
#     stats_data = []
#     stats_data.append(week_titles)
#     for i in list(results.keys()):
#         name_data.append([i])
#         stats_data.append(results[i])
#     stats.update_values(crange = data_range.range, values = stats_data, majordim = 'COLUMNS', parse=True)
#     stats.update_values(crange = name_range.range, values = name_data, majordim = 'COLUMNS', parse=False)

# def get_stats(sheet):
#     cell_list = sheet.range(crange='2:2', returnas='matrix')[0]
#     del cell_list[0:6]
#     names = [i for i in cell_list if i != '']
#     print(names)
#     costs_row = str(sheet.find('Match cost:')[0].row)
#     costs = sheet.range(crange=str(costs_row+':'+costs_row), returnas='matrix')[0]
#     del costs[0:6]
#     costs = [i for i in costs if i != '']
#     print(costs)
#     results = {}
#     for i in range(len(names)):
#         results[names[i]] = costs[i]
#     return results

# backport_stats()

root = Tk()
ui(root)
root.mainloop()