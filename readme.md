A program to calculate science bowl stats based on science bowl scoresheets.

# How to use
First, put all of the scoresheets into one folder, and then place the folder in the same folder that this repository is in. Then, go to the `key.json` file and change the "directory" variable to the relative filepath of the folder. After that, you can run the `stats.py` file to generate the stats.

The python version used is `3.9.1`. You may need to install python dependencies. This can be done by running the following command in the terminal:

```pip install -r requirements.txt```

# JSON Options
The following options can be changed in `key.json`. The default setting for each boolean parameter is `false`.
- directory: stores the filepath to the folder with the scoresheets
- rosters: stores the filepath to the plaintext file with the rosters, which is a comma-seperated file with the name of the player, followed by the name of their team
- category directory: (default: `""`)
    - If every packet has the same category order (e.g. question #1 is always physics, question #2 is always biology), then it may be helpful to include a text file that lists the order of the categories, with each new line denoting a new question category. In that case, this variable would indicate the path to that file.
    - If no such file exists, then set this variable to `""` (an empty string).
- force questions to have categories:
    - If set to `true`, then the program will attempt to detect a category for each question, either by having the category in the column adjacent to the question number column or by having a "subject order" file. __Any question that does NOT have a category will be ignored and not included in the final stat report.__
    - If set to `false`, then which category a question is in will still be tracked, but there is no requirement for questions to have a subject, and questions without a category will still be tracked. 
- skip players with no buzzes:
    - If set to `true`, then don't track stats for players with 0 buzzes.
    - If set to `false`, include stats for *all* players, regardless of how many buzzes they have.
- has interrupt corrects:
    - If set to `true`, then the program will try to track the number of times a player interrupted correctly by looking for an "interrupt correct" symbol in the json file.
    - If set to `false`, then the program does not track this statistic
- track TUH:
    - If set to `true`, then the program will attempt to detect a row on the spreadsheet that has the number of tossups each player heard. 
    - If set to `false`, then the program will assume that each player plays a full game in each spreadsheet that they appear in.
- player names to ignore: 
    - Often times, other strings (like "Bonus" or "Question") are interpreted as player names. Any player names that contain one or more of the strings in this array as substrings will be ignored and assumed to not be a player. 
- cat per packet: stores the number of times each category appears in each packet
- codes: the different cell values that indicate each type of buzz
- categories: the six DOE categories (biology, chemistry, etc.) as well as the different symbols that indicate those categories

# Interpreting the Reports
The report will generate a total of 17 different reports, each one in its own tab. Here are the 17:
- subject, subject_team
    - these two spreadsheets display a simple overview of each individial and team's points per game in each category
- bonus
    - this sheet displays the number of bonuses heard and converted for each team in each category, as well as the % of bonuses converted
- all, all_team, and two reports for each category (individual / team stats)
    - these spreadsheets provide detailed information about each player and team's performance, including the number of correct questions, the number of interrupt incorrects, and points per game. Here are what each of the columns mean:
        - Player/Team: the name of the player/team
        - GP: games played
        - 4I: number of questions where a player/team interrupted a question correctly. Note: this column will not exist if the scoresheets that you use do not support this
        - 4: number of questions a player does not interrupt but gets correct. Note: if "4I" is not supported, then this stat tracks the total number of questions correct, regardless of interrupt
        - -4: number of questions where a player interrupted incorrectly
        - X: number of times where a player buzzed incorrectly *after* the question was completely read.
        - TUH: tossups heard
        - #buzz: the total number of times a player buzzed
        - %buzz: the percentage of tossups heard a player buzzed on
        - %I: the percentage of a player's buzzes that were interrupts
        - 4I/-4: the number of interrupt correct buzzes divided by the number of interrupt incorrect buzzes. Note: this statistic will not appear if "4I" is not supported
        - 4s/-4: the number of correct buzzes (including both interrupts and non-interrupts) divided by the number of interrupt incorrect buzzes
        - P/TUH: the avergae number of points a player gets per tossup they hear
        - Pts: the total number of points the player has scored, where interrupt incorrect buzzes count for -4 points for that player (as opposed to DOE rules, which give 4 points to the other team)
        - PPG: the average number of points a player scores per game