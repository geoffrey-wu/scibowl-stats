
import csv
import json
import os

import numpy as np
import pandas as pd


with open('key.json') as f:
    json_data = json.load(f)

cats = json_data['subjects'].keys()  # different subject categories

# file directory that contains all of the scoresheets
directory = json_data['directory']

# codes for different buzz results
# e.g. interrupt incorrect, interrupt correct
codes = json_data['codes']


# converts a pandas dataframe to a np.array
def to_nparray(dataframe):
    return np.append(np.array([dataframe.columns]), dataframe.to_numpy(), axis=0)


# given a string, returns the question category
def get_category(category):
    category = category.lower().strip()
    for cat in cats:
        if category in json_data['subjects'][cat]:
            return cat
    return 'n/a'


# given a cell value, returns the correct index for the code
def get_code_index(cell):
    cell = str(cell).upper().strip()
    if cell in codes['interrupt_correct']:
        return 0
    elif cell in codes['correct']:
        return 1
    elif cell in codes['neg']:
        return 2
    elif cell in codes['incorrect1']:
        return 3
    elif cell in codes['incorrect2']:
        return 4
    return -1


# returns the row number in the spreadsheet where the questions begin
def get_question_row(game):
    for i in range(game.shape[0]):
        if game[i][0] == 1:
            return i
    return -1


rosters = {}
for line in open(json_data['rosters'], 'r'):
    player, team = line.split(',')
    player = player.strip().title()
    team = team.strip().title()
    rosters[player] = team


# dictionary containing per player stats
per_player_stats = {}

# dictionary containing per team bonus stats
per_team_bonus_stats = {}
for team in rosters.values():
    if team not in per_team_bonus_stats:
        per_team_bonus_stats[team] = {
            'GP': 0
        }
        for cat in cats:
            per_team_bonus_stats[team][cat] = [0, 0]


# reads all of the spreadsheets in the given folder and reads the data in each of the files
for filename in os.listdir(directory):
    print(filename)
    for game in pd.read_excel(directory + '\\' + filename, sheet_name=None).values():
        game = to_nparray(game)
        teams_in_game = []
        question_row = get_question_row(game)

        for j in 2 + np.array(range(game.shape[1] - 2)):
            player = str(game[question_row - 1, j]).title().strip()
            if player in ['Nan', '']:  # ignore if the cell is empty
                continue
            
            if player[:6] == 'Player' or player[:5] == 'Bonus' or player[:11] == 'Total Score' or player[:7] == 'Unnamed' or player == 'Score':
                continue

            # create a new player if the player isn't already in the database
            if player not in per_player_stats:
                per_player_stats[player] = {
                    'GP': 0
                }
                for cat in cats:
                    per_player_stats[player][cat] = len(codes)*[0]

            # if the player is on a team, then add that
            # player's team to the list of teams in the game
            if player in rosters:
                team = rosters[player]
                if team not in teams_in_game:
                    teams_in_game.append(team)

            per_player_stats[player]['GP'] += 1

            # for each player, look down their respective column
            # to collect data on when they buzzed
            for i in question_row + np.array(range(game.shape[0] - question_row)):
                # cell that records their buzz
                cell = str(game[i, j]).upper().strip()
                # category the question was in
                cat = get_category(str(game[i, 1]))
                index = get_code_index(cell)

                if index != -1 and cat != 'n/a':
                    per_player_stats[player]['all'][index] += 1
                    per_player_stats[player][cat][index] += 1
                # if index != -1:
                #     per_player_stats[player]['all'][index] += 1

        # skip if there are fewer than 2 teams
        if len(teams_in_game) < 2:
            continue

        n = 0
        for j in range(game.shape[1]):
            if str(game[1, j]).strip().lower()[:5] == 'bonus' or str(game[0, j]).strip().lower()[:5] == 'bonus':
                for i in 2 + np.array(range(game.shape[0] - 2)):
                    if str(game[i, j]).strip() in ['1', '1.0', '10', '10.0']:
                        cat = get_category(str(game[i, 1]))
                        if cat == 'n/a':
                            continue
                        per_team_bonus_stats[teams_in_game[n]]['all'][0] += 1
                        per_team_bonus_stats[teams_in_game[n]][cat][0] += 1
                n += 1
        per_team_bonus_stats[teams_in_game[0]]['GP'] += 1
        per_team_bonus_stats[teams_in_game[1]]['GP'] += 1


# dictionary containing per category stats
per_cat_stats = {}

# dictionary containing per_team_stats
per_team_tu_stats = {}

# initialize columns of the spreadsheet--indicates type of data included
for cat in cats:
    temp = [[
        'Player',   # name of player
        'GP',       # games played
        '4I',       # interrupt correct
        '4',        # correct (but no interrupt)
        '-4',       # interrupt incorrect
        'X1',       # not interrupt, first wrong buzz
        'X2',       # not interrupt, second wrong buzz
        'TUH',      # tossups heard
        '#buzz',    # number of total buzzes
        '%buzz',    # percent of tossups heard that the player buzzed
        '%I',       # percent of tossups heard that the player interrupted
        '4I/-4',    # interrupt corrects per neg
        '4s/-4',    # corrects per neg
        'P/TUH',    # points per tossup heard
        'Pts',      # total points
        'PPG'       # points per game
    ]]
    per_cat_stats[cat] = [[i for i in temp[0]]]
    per_team_tu_stats[cat] = [[i for i in temp[0]]]

team_to_number = {}
i = 0
for team in rosters.values():
    i += 1
    team_to_number[team] = i
    for cat in cats:
        array = [0]*16
        array[0] = team
        per_team_tu_stats[cat].append(array)


def player_to_team_num(player):
    if player in rosters:
        return team_to_number[rosters[player]]
    else:
        return -1


# compiles per-category stats from the per-player stats
for player in per_player_stats:
    # if a player has no stats, don't include them in the stat report
    # if sum(per_player_stats[player]['all']) == 0:
    #     continue

    GP = per_player_stats[player]['GP']  # games played
    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]

        # TUH = tossups heard
        # this dictionary gives the number of tossups in each category per game
        TUH = GP * {
            'all': 23,
            'bio': 4,
            'chem': 4,
            'energy': 3,
            'ess': 4,
            'math': 4,
            'physics': 4
        }[cat]

        # number of times the player buzzed
        num_buzz = fourI + four + neg + x1 + x2

        # percentage of tossups heard that the player buzzed on
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
        P_TUH = round(points/TUH, 2)         # points per tossup
        ppg = round(points/GP, 2)           # points per game

        per_cat_stats[cat].append([
            player,
            GP,
            fourI,
            four,
            neg,
            x1,
            x2,
            TUH,
            num_buzz,
            pct_buzz,
            pct_I,
            fourI_neg,
            four_neg,
            P_TUH,
            points,
            ppg
        ])

        if player_to_team_num(player) < 0:
            continue
        
        array = per_team_tu_stats[cat][player_to_team_num(player)]  # player
        array[1] = max([array[1], GP])  # GP
        array[2] += fourI  # fourI
        array[3] += four  # four
        array[4] += neg  # neg
        array[5] += x1  # X1
        array[6] += x2  # x2
        array[7] = max(array[7], TUH)  # TUH
        array[8] += num_buzz  # number of buzzes
        array[9] = str(round(100*array[8]/array[7], 2)) + '%' # pct_buzz
        array[10] = 0  # pct_I
        array[11] = 0  # fourI_neg
        array[12] = 'inf' if array[4] == 0 else round(array[3]/array[4], 2)  # four_neg
        array[14] += points  # points
        array[13] = round(array[14]/array[7], 2)  # P_TUH
        array[15] = round(array[14] / array[1], 2)  # ppg

        # per_team_tu_stats[cat][player_to_team_num(player)] = array

# compiles subject and bonus stats
aggregate_subject = [[
    'Player',
    'GP',
    'ppg',
    'bio',
    'chem',
    'energy',
    'ess',
    'math',
    'physics'
]]

aggregate_subject_team = [[
    'Player',
    'GP',
    'ppg',
    'bio',
    'chem',
    'energy',
    'ess',
    'math',
    'physics'
]]

for i in range(len(rosters.values())):
    aggregate_subject_team.append([0]*9)

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
    array2 = aggregate_subject_team[player_to_team_num(player)]
    array2[0] = team
    array2[1] = max(array2[1], GP)
    
    for i in range(7):
        array2[i+2] = round(array2[i+2] + array[i+2], 2)
    
    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]
        per_team_bonus_stats[team][cat][1] += fourI + four  # number of bonuses heard


# converts per_team_bonus_stats, a dictionary, to per_team_stats_array,
# a 2D array used for saving to an excel spreadsheet
per_team_stats_array = [['team', 'GP']]
for cat in cats:
    per_team_stats_array[0].append(cat)
    per_team_stats_array[0].append('')
    per_team_stats_array[0].append('%')

for team in per_team_bonus_stats:
    array = [team, per_team_bonus_stats[team]['GP']]
    for cat in cats:
        correct = per_team_bonus_stats[team][cat][0]
        total = per_team_bonus_stats[team][cat][1]
        percentage = round(100*correct/total, 2)
        array.append(correct)
        array.append(total)
        array.append(percentage)
    per_team_stats_array.append(array)

total_4I = 0
for player in per_player_stats:
    total_4I += per_player_stats[player]['all'][0]

if total_4I == 0:
    for cat in cats:
        per_cat_stats[cat] = np.delete(per_cat_stats[cat], [2, 6, 10, 11], axis=1)
        per_team_tu_stats[cat] = np.delete(per_team_tu_stats[cat], [2, 6, 10, 11], axis=1)


for cat in cats:
    array = []
    for i in range(len(per_team_tu_stats[cat])):
        if per_team_tu_stats[cat][i, 1] in [0, '0', 0.0, '0.0']:
            array.append(i)
    
    per_team_tu_stats[cat] = np.delete(per_team_tu_stats[cat], array, axis=0)

array = []
for i in range(len(aggregate_subject_team)):
    if aggregate_subject_team[i][1] in [0, '0', 0.0, '0.0']:
        array.append(i)
    
aggregate_subject_team = np.delete(aggregate_subject_team, array, axis=0)

# write all the data into spreadsheets
with pd.ExcelWriter(directory + '_stats.xlsx') as writer:
    stat_sheet = pd.DataFrame(np.array(aggregate_subject))
    stat_sheet.to_excel(writer, sheet_name='subject', header=None, index=False)

    stat_sheet = pd.DataFrame(np.array(aggregate_subject_team))
    stat_sheet.to_excel(writer, sheet_name='subject_team', header=None, index=False)

    stat_sheet = pd.DataFrame(np.array(per_team_stats_array))
    stat_sheet.to_excel(writer, sheet_name='bonus', header=None, index=False)

    for cat in cats:
        stat_sheet = pd.DataFrame(np.array(per_cat_stats[cat]))
        stat_sheet.to_excel(writer, sheet_name=cat, header=None, index=False)
        
        stat_sheet = pd.DataFrame(np.array(per_team_tu_stats[cat]))
        stat_sheet.to_excel(writer, sheet_name=cat + '_team', header=None, index=False)
