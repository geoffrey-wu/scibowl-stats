

A program to calculate science bowl stats based on science bowl scoresheets.

# How to use
First, add the folder containing the scoresheets (which must be in microsoft excel files) to this directory. Update the "directory" variable in ``key.json`` to the path of the score sheets (note: this path can be relative or absolute). 

The file ``stats.py`` contains all of the code required to generate the spreadsheets (simply run the file to generate the stat reports). The stat reports will appear in a single Microsoft Excel spreadsheet. 

# Options
The ``key.json`` file provides a tremendous amount of flexibility and adaptability for the user. The following options can be changed in ``key.json``:
- directory: stores the filepath to the folder with the scoresheets
- rosters: stores the filepath to the plaintext file with the rosters, which is a comma-seperated file with the name of the player, followed by the name of their team

# Interpreting the Reports
The report will generate a total of 17 different reports, each one in its own tab. Here are the 17:
- subject, subject_team
    - these two spreadsheets display a simple overview of each individial and team's points per game in each category
- bonus
    - this sheet displays the number of bonuses heard and converted for each team in each category, as well as the % of bonuses converted
- all, all_team, and two reports for each category (individual / team stats)
    - these spreadsheets provide detailed information about each player and team's performance, including the number of correct questions, the number of interrupt incorrects, and points per game. Here are what each of the columns mean:
        - Player: the name of the player/team
        - GP: games played
        - 4I: number of questions where a player/team interrupted a question correctly. Note: this column will not exist if the scoresheets that you use do not support this
        - 4: number of questions a player does not interrupt but gets correct. Note: if "4I" is not supported, then this stat tracks the total number of questions correct, regardless of interrupt
        - -4: number of questions where a player interrupted incorrectly
        - X1: number of times where a player was the first person to buzz incorrectly *after* the question was completely read. Note: if "4I" is not supported, then neither is "X2", and therefore this stat will track the total number of non-interrupt incorrect buzzes, regardless of the players that buzzed before.
        - X2: number of times where a player was the second person to buzz incorrectly *after* the question was completely read. Note: if "4I" is not supported, then neither is this statistic
        - TUH: tossups heard
        - #buzz: the total number of times a player buzzed
        - %buzz: the percentage of tossups heard a player buzzed on
        - #I: the percentage of a player's buzzes that were interrupts
        - 4I/-4: the number of interrupt correct buzzes divided by the number of interrupt incorrect buzzes. Note: this statistic will not appear if "4I" is not supported
        - 4s/-4: the number of correct buzzes (including both interrupts and non-interrupts) divided by the number of interrupt incorrect buzzes
        - P/TUH: the avergae number of points a player gets per tossup they hear
        - pts: the total number of points the player has scored, where interrupt incorrect buzzes count for -4 points for that player (as opposed to DOE roles, which give 4 points to the other team)
        - ppg: the average number of points a player scores per game