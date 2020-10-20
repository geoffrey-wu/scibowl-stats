
import csv
import numpy as np
import os
import pandas as pd
import json


def to_nparray(dataframe):
    return np.append(np.array([dataframe.columns]), dataframe.to_numpy(), axis=0)


def get_category(category):
    category = category.lower()
    if category in ['bio', 'biology', 'b']:
        return cats[1]
    elif category in ['chem', 'chemistry', 'c', 'ch']:
        return cats[2]
    elif category in ['energy', 'en']:
        return cats[3]
    elif category in ['earth and space', 'es', 'ess']:
        return cats[4]
    elif category in ['math', 'm']:
        return cats[5]
    elif category in ['physics', 'phys', 'p']:
        return cats[6]
    else:
        return 'n/a'


# working directory that contains all of the scoresheets
directory = '2020-21 season\\Discord'

with open('key.json') as f:
    json_data = json.load(f)

cats = json_data['cats']
codes = json_data['codes']
all_players = {}

for filename in os.listdir(directory):
    for game in pd.read_excel(directory + '\\' + filename, sheet_name=None).values():
        game = to_nparray(game)

        for j in map(lambda a: a + 2, range(game.shape[1] - 2)):
            player = str(game[1, j]).upper()
            if player == 'NAN':
                continue

            if player not in all_players:
                all_players[player] = {
                    'GP': 0,
                    'n/a': len(codes)*[0],
                }

                for i in range(len(cats)):
                    all_players[player][cats[i]] = len(codes)*[0]

            fourI = four = neg = 0

            for i in map(lambda a: a + 2, range(game.shape[0] - 2)):
                cell = str(game[i, j]).upper()
                cat = get_category(str(game[i, 1]))
                index = -1

                if cell == codes['interrupt_correct']:
                    index = 0
                elif cell == codes['correct']:
                    index = 1
                elif cell == codes['neg']:
                    index = 2
                elif cell == codes['incorrect1']:
                    index = 3
                elif cell == codes['incorrect2']:
                    index = 4

                if index != -1:
                    all_players[player]['all'][index] += 1
                    all_players[player][cat][index] += 1

            all_players[player]['GP'] += 1

stats = {}
for cat in cats:
    stats[cat] = [['Player', 'GP', '4I', '4', '-4', 'X1', 'X2',
                   'TUH', '#buzz', '%buzz', '%I', '4I/-4', '4s/-4', 'P/TU', 'Pts', 'PPG']]

for player in all_players.keys():
    if sum(all_players[player]['all']) == 0:
        continue

    GP = all_players[player]['GP']
    for cat in cats:
        TUH = GP * (23 if cat == 'all' else 4)
        fourI, four, neg, x1, x2 = all_players[player][cat]

        points = 4*fourI + 4*four - 4*neg
        num_buzz = fourI + four + neg + x1 + x2
        pct_buzz = str(round(100*num_buzz/TUH, 2)) + '%'
        P_TU = round(points/TUH, 2)
        if num_buzz != 0:
            pct_interrupt = str(round(100*(fourI + neg)/num_buzz, 2)) + '%'
        else:
            pct_interrupt = 'N/A'

        if neg == 0:
            fourI_neg = 0 if fourI == 0 else 'inf'
            four_neg = 0 if four + fourI == 0 else 'inf'
        else:
            fourI_neg = round(fourI/neg, 2)
            four_neg = round((fourI + four)/neg, 2)

        ppg = round(points/GP, 2)
        stats[cat].append([player, GP, fourI, four, neg, x1, x2, TUH, num_buzz,
                           pct_buzz, pct_interrupt, fourI_neg, four_neg, P_TU, points, ppg])

for cat in cats:
    filename = directory + '_stats_' + cat + '.csv'
    stat_sheet = pd.DataFrame(np.array(stats[cat]))
    stat_sheet.to_csv(filename, header=None, index=False)
