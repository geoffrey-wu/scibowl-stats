
import csv
import json
import os

import numpy as np
import pandas as pd


# converts a pandas dataframe to a np.array
def to_nparray(dataframe):
    return np.append(np.array([dataframe.columns]), dataframe.to_numpy(), axis=0)


# given a string, returns the question category
def get_category(category):
    category = category.lower().strip()
    if category in ['b', 'bio', 'biology']:
        return cats[1]
    elif category in ['c', 'ch', 'chem', 'chemistry']:
        return cats[2]
    elif category in ['en', 'energy']:
        return cats[3]
    elif category in ['earth and space', 'es', 'ess']:
        return cats[4]
    elif category in ['m', 'math']:
        return cats[5]
    elif category in ['p', 'ph', 'phy', 'phys', 'physics']:
        return cats[6]
    else:
        return 'n/a'


with open('key.json') as f:
    json_data = json.load(f)

# different subject categories
cats = ['all', 'bio', 'chem', 'energy', 'ess', 'math', 'phys']

# file directory that contains all of the scoresheets
directory = json_data['directory']

# codes for different buzz results
# e.g. interrupt incorrect, interrupt correct
codes = json_data['codes']

# dictionary containing per player stats
per_player_stats = {}

rosters = {}
for line in open('rosters.txt', 'r'):
    player, team = line.split(',')
    player = player.strip().upper()
    team = team.strip().upper()
    
    rosters[player] = team

per_team_stats = {}
for team in rosters.values():
    if team not in per_team_stats:
        per_team_stats[team] = {
            'GP': 0
        }
        for cat in cats:
            per_team_stats[team][cat] = [0, 0]

for filename in os.listdir(directory):
    print(filename)
    for game in pd.read_excel(directory + '\\' + filename, sheet_name=None).values():
        game = to_nparray(game)
        teams = []

        loc_of_1 = 0  # location of the first one in the question number column
        for i in range(game.shape[0]):
            if game[i][0] == 1:
                loc_of_1 = i
                break

        for j in range(game.shape[1] - 2):
            player = str(game[loc_of_1 - 1, j + 2]).upper().strip()
            # ignore if the cell is empty
            if player in ['NAN', '']:
                continue

            # create a new player if the player isn't already in the database
            if player not in per_player_stats:
                per_player_stats[player] = {
                    'GP': 0
                }
                for cat in cats:
                    per_player_stats[player][cat] = len(codes)*[0]

            if player in rosters:
                team = rosters[player]
                if team not in teams:
                    teams.append(team)

            for i in range(game.shape[0] - loc_of_1):
                cell = str(game[i + loc_of_1, j + 2]).upper().strip()
                cat = get_category(str(game[i + loc_of_1, 1]))
                index = -1

                if cell in codes['interrupt_correct']:
                    index = 0
                elif cell in codes['correct']:
                    index = 1
                elif cell in codes['neg']:
                    index = 2
                elif cell in codes['incorrect1']:
                    index = 3
                elif cell in codes['incorrect2']:
                    index = 4

                if index != -1 and cat != 'n/a':
                    per_player_stats[player]['all'][index] += 1
                    per_player_stats[player][cat][index] += 1

            per_player_stats[player]['GP'] += 1

        if len(teams) != 2:
            continue

        team_number = 0
        for j in range(game.shape[1]):
            if str(game[1, j]).strip().lower() == 'bonus' or str(game[0, j]).strip().lower() == 'bonus':
                for i in range(game.shape[0] - 2):
                    if str(game[i + 2, j]).strip() in ['1', '1.0', '10', '10.0']:
                        cat = get_category(str(game[i + 2, 1]))
                        
                        if cat != 'n/a':
                            per_team_stats[teams[team_number]]['all'][0] += 1
                            per_team_stats[teams[team_number]][cat][0] += 1
                team_number += 1
        per_team_stats[teams[0]]['GP'] += 1
        per_team_stats[teams[1]]['GP'] += 1


# dictionary containing per category stats
per_cat_stats = {}

# initialize columns of the spreadsheet--indicates type of data included
for cat in cats:
    per_cat_stats[cat] = [['Player', 'GP', '4I', '4', '-4', 'X1', 'X2',
                           'TUH', '#buzz', '%buzz', '%I', '4I/-4', '4s/-4', 'P/TU', 'Pts', 'PPG']]

for player in per_player_stats.keys():
    # if a player has no stats, don't include them in the stat report
    if sum(per_player_stats[player]['all']) == 0:
        continue

    GP = per_player_stats[player]['GP']
    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]

        # tossups heard
        TUH = GP*(23 if cat == 'all' else 4)

        # number of times the player buzzed
        num_buzz = fourI + four + neg + x1 + x2

        # percentage of tossups that the player buzzed on
        pct_buzz = str(round(100*num_buzz/TUH, 2)) + '%'

        # percentage of buzzes that are an interrupt
        if num_buzz != 0:
            pct_I = str(round(100*(fourI + neg)/num_buzz, 2)) + '%'
        else:
            pct_I = 'N/A'

        if neg == 0:
            fourI_neg = 0 if fourI == 0 else 'inf'
            four_neg = 0 if four + fourI == 0 else 'inf'
        else:
            fourI_neg = round(fourI/neg, 2)
            four_neg = round((fourI + four)/neg, 2)

        points = 4*fourI + 4*four - 4*neg   # total points scored
        P_TU = round(points/TUH, 2)         # points per tossup
        ppg = round(points/GP, 2)           # points per game

        per_cat_stats[cat].append([player, GP, fourI, four, neg, x1, x2, TUH, num_buzz,
                                   pct_buzz, pct_I, fourI_neg, four_neg, P_TU, points, ppg])

aggregate_subject = [['Player', 'GP', 'ppg', 'bio',
                      'chem', 'energy', 'ess', 'math', 'phys']]
for player in per_player_stats:
    GP = per_player_stats[player]['GP']
    array = [player, GP]
    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]
        array.append(round(4*(fourI + four - neg)/GP, 2))
    aggregate_subject.append(array)

    if player not in rosters:
        continue
    team = rosters[player]
    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]
        per_team_stats[team][cat][1] += fourI + four

# write all the subject data into spreadsheets
# with pd.ExcelWriter(directory + '_stats_noyes.xlsx') as writer:
#     stat_sheet = pd.DataFrame(np.array(aggregate_subject))
#     stat_sheet.to_excel(writer, sheet_name='subject', header=None, index=False)
#     for cat in cats:
#         stat_sheet = pd.DataFrame(np.array(per_cat_stats[cat]))
#         stat_sheet.to_excel(writer, sheet_name=cat, header=None, index=False)

for team in per_team_stats:
    print(team)
    print(per_team_stats[team])
