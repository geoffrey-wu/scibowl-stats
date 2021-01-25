
import csv 
import os
import pandas as pd
import re
import random 
import numpy as np
import json

directory = 'Results'

with open('key.json') as f:
    json_data = json.load(f)

# returns the row number in the spreadsheet where the questions begin
def get_question_row(game):
    for i in range(game.shape[0]):
        if game[i][0] == 1:
            return i
    return -1

def write_to_excel(writer, data, name):
    stat_sheet = pd.DataFrame(np.array(data))
    stat_sheet.to_excel(writer, sheet_name=name, header=None, index=False)

for (dirpath, dirnames, filenames) in os.walk(directory):
    for filename in filenames:
        filepath = dirpath + '/' + filename
        print(filepath)
        if filepath[-9:] == '.DS_Store':
            continue
        game = 0
        sheetname = 'Sheet1'
        for sheetname in pd.read_excel(filepath, sheet_name=None):
            # print(sheetname)
            game = pd.read_excel(filepath, sheet_name=None)[sheetname]
            # converts the pandas dataframe to a np.array
            game = np.append(np.array([game.columns]), game.to_numpy(), axis=0)
            question_row = get_question_row(game)

            for j in 1 + np.array(range(game.shape[1] - 1)):
                player = str(game[question_row - 1, j]).title().strip()
                skip_player = player in ['', 'Nan']
                for s in json_data['player names to ignore']:
                    skip_player = skip_player or (s in player)

                if skip_player:
                    continue
                
                player = re.sub('\[[^\]]*\]', '', player).strip()
                game[question_row - 1, j] = player

        with pd.ExcelWriter(filepath) as writer:
            write_to_excel(writer, game, sheetname)


