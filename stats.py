
import csv
import numpy as np
import os
import pandas as pd


def convert_to_nparray(dataframe):
    return np.append(np.array([dataframe.columns]), dataframe.to_numpy(), axis=0)

# changable params
directory = '2018-19 season'
interrupt_correct = '4I'
correct = '4'
neg = 'XX'

all_players = {}

for filename in os.listdir(directory):
    for read_file in pd.read_excel(directory + '\\' + filename, sheet_name=None).values():
        read_file = convert_to_nparray(read_file)

        for j in map(lambda a: a + 2, range(read_file.shape[1] - 2)):
            name = read_file[1, j]
            name = str(name).upper()
            if name == 'NAN':
                continue

            fourI, four, neg = 0, 0, 0

            for i in map(lambda a: a + 2, range(read_file.shape[0] - 2)):
                cell = str(read_file[i, j]).upper()
                if cell == '4I':
                    fourI += 1
                elif cell == '4':
                    four += 1
                elif cell == 'XX':
                    neg += 1

            if name in all_players:
                all_players[name]['fourI'] += fourI
                all_players[name]['four'] += four
                all_players[name]['neg'] += neg
                all_players[name]['GP'] += 1
            else:
                all_players[name] = {
                    'fourI': fourI,
                    'four': four,
                    'neg': neg,
                    'GP': 1
                }

# print(all_players)

stats = [['Player', 'GP', '4I', '4', '-4', 'TUH',
          'P/TU', '4I/-4', '4s/-4', 'Pts', 'PPG']]

for name in all_players.keys():
    GP = all_players[name]['GP']
    fourI = all_players[name]['fourI']
    four = all_players[name]['four']
    neg = all_players[name]['neg']
    TUH = GP * 23
    points = 4*fourI + 4*four - 4*neg
    P_TU = round(points / TUH, 2)

    fourI_neg, four_neg = 0, 0
    if neg == 0:
        fourI_neg = 0 if fourI == 0 else 'inf'
        four_neg = 0 if four + fourI == 0 else 'inf'
    else:
        fourI_neg = round(fourI/neg, 2)
        four_neg = round((fourI + four)/neg, 2)

    ppg = round(points/GP, 2)
    stats.append([name, GP, fourI, four, neg, TUH, P_TU, fourI_neg, four_neg, points, ppg])

stats = np.array(stats)
print(pd.DataFrame(stats))
pd.DataFrame(stats).to_csv(directory + '_stats.csv', header=None, index=False)
