
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

        # converts the pandas dataframe to a np.array
        game = np.append(np.array([game.columns]), game.to_numpy(), axis=0)
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

                if index != -1:
                    if not json_data['force questions to have categories'] or cat != 'n/a':
                        per_player_stats[player]['all'][index] += 1
                        per_player_stats[player][cat][index] += 1

        # skip if there are fewer than 2 teams
        if len(teams_in_game) < 2:
            continue

        n = 0
        for j in range(game.shape[1]):
            if 'bonus' in [str(game[1, j])[:5], str(game[0, j])[:5]].strip().lower():
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
        player_data = [0]*16
        player_data[0] = team
        per_team_tu_stats[cat].append(player_data)


def player_to_team_num(player):
    if player in rosters:
        return team_to_number[rosters[player]]
    else:
        return -1


# compiles per-category stats from the per-player stats
for player in per_player_stats:
    # if a player has no stats, don't include them in the stat report
    if json_data['skip players with no buzzes'] and sum(per_player_stats[player]['all']) == 0:
        continue

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
        P_TUH = round(points/TUH, 2)        # points per tossup
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

        player_data = per_team_tu_stats[cat][player_to_team_num(
            player)]  # player
        player_data[1] = max([player_data[1], GP])  # GP
        player_data[2] += fourI  # fourI
        player_data[3] += four  # four
        player_data[4] += neg  # neg
        player_data[5] += x1  # X1
        player_data[6] += x2  # x2
        player_data[7] = max(player_data[7], TUH)  # TUH
        player_data[8] += num_buzz  # number of buzzes
        player_data[9] = str(
            round(100*player_data[8]/player_data[7], 2)) + '%'  # pct_buzz
        player_data[10] = 0  # pct_I
        player_data[11] = 0  # fourI_neg
        player_data[12] = 'inf' if player_data[4] == 0 else round(
            player_data[3]/player_data[4], 2)  # four_neg
        player_data[14] += points  # points
        player_data[13] = round(player_data[14]/player_data[7], 2)  # P_TUH
        player_data[15] = round(player_data[14] / player_data[1], 2)  # ppg

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
    player_data = [player, GP]
    for cat in cats:  # append the points per game for that category
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]
        player_data.append(round(4*(fourI + four - neg)/GP, 2))

    aggregate_subject.append(player_data)

    if player not in rosters:
        continue

    team = rosters[player]

    array2 = aggregate_subject_team[player_to_team_num(player)]
    array2[0] = team
    array2[1] = max(array2[1], GP)

    for i in range(7):
        array2[i+2] = round(array2[i+2] + player_data[i+2], 2)

    for cat in cats:
        fourI, four, neg, x1, x2 = per_player_stats[player][cat]

        # number of bonuses heard
        per_team_bonus_stats[team][cat][1] += fourI + four


# converts per_team_bonus_stats, a dictionary, to per_team_stats_array,
# a 2D array used for saving to an excel spreadsheet
per_team_stats_array = [['team', 'GP']]
for cat in cats:
    per_team_stats_array[0].append(cat)
    per_team_stats_array[0].append('')
    per_team_stats_array[0].append('%')

for team in per_team_bonus_stats:
    player_data = [team, per_team_bonus_stats[team]['GP']]
    for cat in cats:
        correct = per_team_bonus_stats[team][cat][0]
        total = per_team_bonus_stats[team][cat][1]
        percentage = 0 if total == 0 else round(100*correct/total, 2)
        player_data.append(correct)
        player_data.append(total)
        player_data.append(percentage)
    per_team_stats_array.append(player_data)


# if the spreadsheets do NOT have designations for 
# tossups that were interrupted correctly, then
# delete all columns which rely on this data
if json_data['has interrupt corrects'] == False:
    for cat in cats:
        per_cat_stats[cat] = np.delete(
            per_cat_stats[cat], [2, 6, 10, 11], axis=1)
        per_team_tu_stats[cat] = np.delete(
            per_team_tu_stats[cat], [2, 6, 10, 11], axis=1)


def delete_empty_rows(array2):
    rows_to_delete = []
    for i in range(len(array2)):
        if array2[i][1] in [0, '0', 0.0, '0.0']:
            rows_to_delete.append(i)

    return np.delete(array2, rows_to_delete, axis=0)

aggregate_subject_team = delete_empty_rows(aggregate_subject_team)
for cat in cats:
    per_team_tu_stats[cat] = delete_empty_rows(per_team_tu_stats[cat])


def write_to_excel(writer, data, name):
    stat_sheet = pd.DataFrame(np.array(aggregate_subject))
    stat_sheet.to_excel(writer, sheet_name=name, header=None, index=False)


# write all the data into spreadsheets
with pd.ExcelWriter(directory + '_stats.xlsx') as writer:
    write_to_excel(writer, aggregate_subject, 'subject')
    write_to_excel(writer, aggregate_subject_team, 'subject_team')
    write_to_excel(writer, per_team_stats_array, 'bonus')
    for cat in cats:
        write_to_excel(writer, per_cat_stats[cat], cat)
        write_to_excel(writer, per_team_tu_stats[cat], cat + '_team')
