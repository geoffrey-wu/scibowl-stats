import csv
import json
import os
import re

import numpy as np
import pandas as pd


# computes the Levenshtein distance between two strings a and b
def ldist(a, b):
    size_x = len(a) + 1
    size_y = len(b) + 1
    matrix = np.zeros((size_x, size_y))
    for x in range(size_x):
        matrix[x, 0] = x
    for y in range(size_y):
        matrix[0, y] = y

    for x in range(1, size_x):
        for y in range(1, size_y):
            if a[x-1] == b[y-1]:
                matrix[x, y] = min(
                    matrix[x-1, y] + 1,
                    matrix[x-1, y-1],
                    matrix[x, y-1] + 1
                )
            else:
                matrix[x, y] = min(
                    matrix[x-1, y] + 1,
                    matrix[x-1, y-1] + 1,
                    matrix[x, y-1] + 1
                )

    return (matrix[size_x - 1, size_y - 1])


# finds the element in `array` that best matches a given string `s` based on Levenshtein distance
def find_closest_match(s, array):
    bestWord = ''
    bestDistance = 1000000

    for string in array:
        if (ldist(s, string) < bestDistance):
            bestDistance = ldist(s, string)
            bestWord = string

    return bestWord


# initialize variables from json file
with open('key.json') as f:
    json_data = json.load(f)

    level = 'HS cats' if json_data['is high school'] else 'MS cats'

    # different subject categories
    cats = [cat for cat in json_data[level]['categories'].keys()]
    # file directory that contains all of the scoresheets
    directory = json_data['directory']

    # codes for different buzz results
    # e.g. interrupt incorrect, interrupt correct
    codes = json_data['codes']


# given a string, returns the question category
def get_category(category):
    category = str(category).lower().strip()
    for cat in cats:
        if category in json_data[level]['categories'][cat]:
            return cat
    return 'n/a'


# given a cell value, returns the correct index for the code
def get_code_index(cell):
    cell = str(cell).lower().strip()
    if cell in codes['interrupt_correct']:
        return 0
    elif cell in codes['correct']:
        return 1
    elif cell in codes['interrupt_incorrect']:
        return 2
    elif cell in codes['incorrect']:
        return 3
    return -1


# returns the row number in the spreadsheet where the questions begin
def get_question_row(game):
    for i in range(game.shape[0]):
        if game[i][0] == 1:
            return i
    return -1


### read all of the players in the rosters file ###
rosters = {}
if json_data['rosters'] != '':
    for line in open(json_data['rosters'], 'r'):
        player, team = line.split(',')
        player = player.strip().title()
        team = team.strip().title()
        rosters[player] = team

teams = [team for team in set(rosters.values())]

# dictionary containing per-player stats
player_stats = {}

# dictionary containing per-team bonus stats
bonus_stats = {}
for team in teams:
    bonus_stats[team] = {
        'GP': 0
    }
    for cat in cats:
        bonus_stats[team][cat] = [0, 0]

# reads all of the spreadsheets in the given folder and subfolders and reads the data in each of the files
for (dirpath, dirnames, filenames) in os.walk(directory):
    for filename in filenames:
        filepath = dirpath + '/' + filename
        print(filepath)

        # skip files that do not have an excel file extension
        if filepath[-5:] != '.xlsx':
            continue

        for game in pd.read_excel(filepath, sheet_name=None).values():
            # converts the pandas dataframe to a np.array
            game = np.append(np.array([game.columns]), game.to_numpy(), axis=0)

            teams_in_game = []  # contains the list of teams in the game
            question_row = get_question_row(game)

            if question_row == -1:
                continue

            for j in 1 + np.array(range(game.shape[1] - 1)):
                player = str(game[question_row - 1, j]).title().strip()

                # skip any player name that isn't valid
                if player in json_data['player names to ignore']:
                    continue

                # regex that removes anything enclosed in brackets [] from a player's name
                # player = re.sub('\[[^\]]*\]', '', player).strip()

                if (json_data['force players onto rosters']):
                    player2 = player
                    player = find_closest_match(player, rosters.keys())
                    if (ldist(player, player2) > 0):
                        print('Changed', player2, 'to', player)

                # create a new player if the player isn't already in the database
                if player not in player_stats:
                    player_stats[player] = {
                        'GP': 0,
                        'TUH': 0
                    }
                    for cat in cats:
                        player_stats[player][cat] = len(codes)*[0]

                # if the player is on a team, then add that
                # player's team to the list of teams in the game
                if player in rosters:
                    team = rosters[player]
                    if team not in teams_in_game:
                        teams_in_game.append(team)

                player_stats[player]['GP'] += 1  # updated games played

                # for each player, look down their respective column
                # to collect data on when they buzzed
                for i in question_row + np.array(range(game.shape[0] - question_row)):
                    # cell that records their buzz
                    cell = str(game[i, j]).upper().strip()

                    # check if we have reached a "tossups heard" cell
                    for string in [game[i, 0], game[i - 1, j]]:
                        if str(string).upper().strip() in 'PLAYER TUH TU HEARD' and cell != 'NAN':
                            player_stats[player]['TUH'] += int(cell)

                    # get the category the question was in
                    if json_data['category directory'] == '':
                        cat = game[i, 1]
                    else:
                        cat = open(json_data['category directory'], 'r').readlines()[
                            (i - question_row) % 23]
                    cat = get_category(cat)

                    # get what type of buzz was recorded (e.g. correct, interrupt)
                    index = get_code_index(cell)

                    # add the buzz to the correct category
                    if index != -1:
                        if cat != 'n/a':
                            player_stats[player]['all'][index] += 1
                            player_stats[player][cat][index] += 1
                        elif not json_data['force questions to have categories']:
                            player_stats[player]['all'][index] += 1

            # skip bonus stats if there are fewer than 2 teams
            if len(teams_in_game) < 2:
                continue

            # find bonus stats
            n = 0
            for j in range(game.shape[1]):
                if 'bonus' in str(game[1, j]).lower() or 'bonus' in str(game[0, j]).lower():
                    for i in range(game.shape[0]):
                        # check if the team got the bonus right
                        if str(game[i, j]).strip() in ['1', '1.0', '10', '10.0']:

                            # get which category the bonus was in
                            if json_data['category directory'] == '':
                                cat = game[i, 1]
                            else:
                                cat = open(json_data['category directory'], 'r').readlines()[
                                    (i - question_row) % 23]
                            cat = get_category(cat)

                            if cat == 'n/a':  # only generate bonus stats if the question has a category
                                continue
                            bonus_stats[teams_in_game[n]]['all'][0] += 1
                            bonus_stats[teams_in_game[n]][cat][0] += 1
                    n += 1
            bonus_stats[teams_in_game[0]]['GP'] += 1
            bonus_stats[teams_in_game[1]]['GP'] += 1


cat_stats = {}      # dictionary containing per-category stats
team_tu_stats = {}  # dictionary containing per-team tossup stats

# initialize columns of the spreadsheet--indicates type of data included
header = [
    'Player',   # name of player
    'GP',       # games played
    '4I',       # interrupt correct
    '4',        # correct (but no interrupt)
    '-4',       # interrupt incorrect
    'X',        # not interrupt, wrong buzz
    'TUH',      # tossups heard
    '#buzz',    # number of total buzzes
    '%buzz',    # percent of tossups heard that the player buzzed
    '%I',       # percent of tossups heard that the player interrupted
    '4I/-4',    # interrupt corrects per neg
    '4s/-4',    # corrects per neg
    'P/TUH',    # points per tossup heard
    'Points',   # total points
    'PPG'       # points per game
]
for cat in cats:
    # initialize these arrays to deep copies of the header
    cat_stats[cat] = [[i for i in header]]
    team_tu_stats[cat] = [[i for i in header]]
    team_tu_stats[cat][0][0] = 'Team'  # Change 'Player' to 'Team'


team_to_number = {}
i = 1
for team in teams:
    team_to_number[team] = i
    i += 1
    for cat in cats:
        team_data = [0]*len(team_tu_stats[cat][0])
        team_data[0] = team
        team_tu_stats[cat].append(team_data)


def player_to_team_num(player):
    return -1 if player not in rosters else team_to_number[rosters[player]]


# compiles per-category stats from the per-player stats
for player in player_stats:
    if json_data['track TUH']:
        TUH_total = player_stats[player]['TUH']   # tossups heard
        GP = round(TUH_total/23, 2)
        player_stats[player]['GP'] = GP
    else:
        GP = player_stats[player]['GP']    # games played
        TUH_total = GP * 23

    if GP == 0:
        continue

    for cat in cats:
        fourI, four, neg, x1 = player_stats[player][cat]

        # TUH = tossups heard
        # this dictionary gives the number of tossups in each category per game
        TUH = round(json_data[level]['per packet'][cat] *
                    (GP if json_data['track TUH'] else TUH_total/23))

        # number of times the player buzzed
        num_buzz = fourI + four + neg + x1

        # percentage of tossups heard that the player buzzed on
        pct_buzz = str(round(100*num_buzz/TUH, 2)) + '%'

        # percentage of buzzes that are an interrupt
        pct_I = str(round(100*(fourI + neg)/num_buzz, 2)) + \
            '%' if num_buzz != 0 else 'N/A'

        if neg == 0:
            fourI_neg = 0 if fourI == 0 else 'inf'
            four_neg = 0 if four + fourI == 0 else 'inf'
        else:
            fourI_neg = round(fourI/neg, 2)
            four_neg = round((fourI + four)/neg, 2)

        points = 4*fourI + 4*four - 4*neg   # total points scored
        P_TUH = round(points/TUH, 2)        # points per tossup
        ppg = round(points/GP, 2)           # points per game

        cat_stats[cat].append([
            player,
            GP,
            fourI,
            four,
            neg,
            x1,
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

        team_data = team_tu_stats[cat][player_to_team_num(player)]  # team
        team_data[1] = max(
            [team_data[1], GP, bonus_stats[rosters[player]]['GP']])  # GP
        team_data[2] += fourI   # fourI
        team_data[3] += four    # four
        team_data[4] += neg     # neg
        team_data[5] += x1      # X1
        team_data[6] = max([team_data[6], TUH, round(
            team_data[1] * json_data[level]['per packet'][cat])])  # TUH
        team_data[7] += num_buzz  # number of buzzes
        team_data[8] = str(
            round(100*team_data[7]/team_data[6], 2)) + '%'  # pct_buzz
        team_data[9] = 0  # pct_I
        team_data[10] = 0  # fourI_neg
        team_data[11] = 'inf' if team_data[4] == 0 else round(
            team_data[3]/team_data[4], 2)  # four_neg
        team_data[13] += points  # points
        team_data[12] = round(team_data[13]/team_data[6], 2)  # P_TUH
        team_data[14] = round(team_data[13]/team_data[1], 2)  # ppg


# count how many tossups each team heard
for cat in cats:
    for i in range(len(teams)):
        bonus_stats[teams[i]][cat][1] += int(
            team_tu_stats[cat][i+1][2]) + int(team_tu_stats[cat][i+1][3])


# compiles subject and bonus stats

# spreadsheet header
if json_data['is high school']:
    header = [
        'Player',
        'GP',
        'ppg',
        'bio',
        'chem',
        'energy',
        'ess',
        'math',
        'physics'
    ]
else:
    header = [
        'Player',
        'GP',
        'ppg',
        'life science',
        'energy',
        'ess',
        'math',
        'physical science'
    ]
aggregate_subject = [[i for i in header]]
aggregate_subject_team = [[i for i in header]]
aggregate_subject_team[0][0] = 'Team'

for i in range(len(teams)):
    aggregate_subject_team.append([0]*len(header))

for player in player_stats:
    GP = player_stats[player]['GP']
    if GP == 0:
        continue
    team_data = [player, GP]
    for cat in cats:  # append the points per game for that category
        fourI, four, neg, x1 = player_stats[player][cat]
        team_data.append(round(4*(fourI + four - neg)/GP, 2))

    aggregate_subject.append(team_data)

for i in range(len(teams)):
    i = i + 1
    array2 = aggregate_subject_team[i]
    array2[0] = team_tu_stats['all'][i][0]  # team name
    array2[1] = team_tu_stats['all'][i][1]  # games played
    for j in range(len(cats)):
        array2[j+2] = team_tu_stats[cats[j]][i][-1]


# converts bonus_stats, a dictionary, to bonus_stats_array,
# a 2D array used for saving to an excel spreadsheet
team_bonus_stats_array = [['Team', 'GP']]
for cat in cats:
    team_bonus_stats_array[0].append(cat)
    team_bonus_stats_array[0].append('')
    team_bonus_stats_array[0].append('%')

for team in bonus_stats:
    team_data = [team, bonus_stats[team]['GP']]
    for cat in cats:
        correct = bonus_stats[team][cat][0]
        total = bonus_stats[team][cat][1]

        # calculate the bonus conversion rate
        percentage = 0 if total == 0 else round(100*correct/total, 2)
        team_data.append(correct)
        team_data.append(total)
        team_data.append(percentage)
    team_bonus_stats_array.append(team_data)


# if the spreadsheets do NOT have designations for
# tossups that were interrupted correctly, then
# delete all columns which rely on this data
if json_data['has interrupt corrects'] == False:
    for cat in cats:
        cat_stats[cat] = np.delete(cat_stats[cat], [2, 9, 10], axis=1)
        team_tu_stats[cat] = np.delete(team_tu_stats[cat], [2, 9, 10], axis=1)


def write_to_excel(writer, data, name):
    stat_sheet = pd.DataFrame(np.array(data))
    stat_sheet.to_excel(writer, sheet_name=name, header=None, index=False)


# write all the data into spreadsheets
with pd.ExcelWriter(directory + '_stats.xlsx') as writer:
    write_to_excel(writer, aggregate_subject, 'subject')
    if json_data['rosters'] != '':
        write_to_excel(writer, aggregate_subject_team, 'subject_team')
        write_to_excel(writer, team_bonus_stats_array, 'bonus')
    for cat in cats:
        write_to_excel(writer, cat_stats[cat], cat)
        if json_data['rosters'] != '':
            write_to_excel(writer, team_tu_stats[cat], cat + '_team')
