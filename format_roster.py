

filepath = 'rosters.txt'
roster = []

for line in open(filepath, 'r'):
    l = line.split('	')
    print(l)
    team = l[0].strip()
    for player in l[1:]:
        if len(player.strip()) == 0:
            continue
        roster.append(player.strip() + ',' + team)

f = open(filepath, 'a')
for line in roster:
    f.write(line + '\n')
