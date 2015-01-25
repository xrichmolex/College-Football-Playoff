#!usr/bin/python

import sys, csv
import xlrd
import pprint
from math import pow

#########################################################################################
# Predicting Teams Best Suited for Participation in the College Football Playoff	#
#											# 
# Written by: Richard E. Molina								#
#											#
#											#
# Description:										#
# This code will read in data from files in the Data/ directory and, utilizing these	#
# statistics, will calculate statistical factors needed to determine the best teams	#
# in college football.									#
#											# 
#########################################################################################

# Data Files
RANKINGS = 'Data/Rankings.xlsx'
STANDINGS = 'Data/Final Standings.xlsx'
OFFENSE = 'Data/Total Team Offense.xlsx'
DEFENSE = 'Data/Total Team Defense.xlsx'
PASSING = 'Data/Team Passing Offense.xlsx'
RUSHING = 'Data/Team Rushing Offense.xlsx'
NAMES = 'Data/Team Names.xlsx'

pp = pprint.PrettyPrinter()

#NOTES
# look into 'Offense Strategy' stat: pass att / rush att
# Also, 'Turnover Differential': Takeaways - Giveaways

# Global Variables
rankings = {}
standings = {}
stats = {}
conferences = {}
espn_names = {}		# Get names for cfb20XXstats files
cfb_names = {}		# Get names for standings/rankings data

TEST_YEAR = '2014'


# Methods
def import_excel(filename):
    wb = xlrd.open_workbook(filename)

    data = {}
    for s in wb.sheets():
        #print 'Sheet:',s.name
	sheet = str(s.name)
	teams = []
        for row in range(s.nrows):
	    values = []
	    for col in range(s.ncols):
	        values.append(s.cell(row,col).value)
	    teams.append(values)
	    #print values
        data[sheet] = teams

    return data

def team_mapping():

    name_map = import_excel(NAMES)
    cfb_names = {}
    espn_names = {}

    for sheet in name_map.keys():
	#print "\n\n"
	#print sheet
	count = 0
	for line in name_map[sheet]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    espn = str(line[0])
	    cfb = str(line[1])

	    cfb_names[espn] = cfb
	    espn_names[cfb] = espn

    return espn_names, cfb_names


def import_stats():

    stats = {}

    # Total Offense
    # [Yards, Yards/Game, Pass Yards, Pass Yards/Game, Rush Yards, Rush Yards/Game, Points, Points/Game]
    rawdata = import_excel(OFFENSE)

    for year in rawdata.keys():
	#print "\n\n"
	#print year
	data = {}
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    rank = int(line[0])
	    team = str(line[1])
	    yds = int(line[2])
	    yds_g = int(line[3])
	    pas = int(line[4])
	    pass_g = float(line[5])
	    rush = int(line[6])
	    rush_g = float(line[7])
	    pts = int(line[8])
	    pts_g = float(line[9])

	    off = {}
	    off['Offense'] = [rank, yds, yds_g, pas, pass_g, rush, rush_g, pts, pts_g]
	    data[team] = off
	    #print team+": "+str(data[team])

	stats[year] = data

    # Total Defense
    # [Yards, Yards/Game, Pass Yards, Pass Yards/Game, Rush Yards, Rush Yards/Game, Points, Points/Game]
    rawdata = import_excel(DEFENSE)

    for year in rawdata.keys():
	#print "\n\n"
	#print year
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    rank = int(line[0])
	    team = str(line[1])
	    yds = int(line[2])
	    yds_g = int(line[3])
	    pas = int(line[4])
	    pass_g = float(line[5])
	    rush = int(line[6])
	    rush_g = float(line[7])
	    pts = int(line[8])
	    pts_g = float(line[9])

	    defen = [rank, yds, yds_g, pas, pass_g, rush, rush_g, pts, pts_g]
	    stats[year][team]['Defense'] = defen
	    #print team+": "+str(stats[year][team])

    # Passing
    # [Attempts, Completions, Percentage, Yards/Att, TDs, INTs, Sacks, Rating]
    rawdata = import_excel(PASSING)

    for year in rawdata.keys():
	#print "\n\n"
	#print year
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    rank = int(line[0])
	    team = str(line[1])
	    att = int(line[2])
	    comp = int(line[3])
	    pct = float(line[4])
	    yds_a = float(line[6])
	    td = int(line[8])
	    inter = int(line[9])
	    sack = int(line[10])
	    rat = float(line[12])

	    passing = [rank,att, comp, pct, yds_a, td, inter, sack, rat]
	    stats[year][team]['Passing'] = passing
	    #print team+": "+str(stats[year][team])

    # Rushing
    # [Attempts, Yards, Yards/Att, TDs]
    rawdata = import_excel(RUSHING)

    for year in rawdata.keys():
	#print "\n\n"
	#print year
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    rank = int(line[0])
	    team = str(line[1])
	    att = int(line[2])
	    yds = int(line[3])
	    yds_a = float(line[4])
	    td = int(line[6])

	    rushing = [rank, att, yds, yds_a, td]
	    stats[year][team]['Rushing'] = rushing
	    #print team+": "+str(stats[year][team])

    return stats

def import_standings():

    global conferences

    standings = {}
    rawdata = import_excel(STANDINGS)

    for year in rawdata.keys():
	print "\n\n"
	print year
	data = {}
	confern = {}
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    conf = str(line[0])
	    team = str(line[1])
	    champ = str(line[2])
	    wins = int(line[7])	
	    loss = int(line[8])
	    perct = "%.3f" % (wins/float(wins+loss))

	    if not conf == 'IND':
		c_win = int(line[3])
		c_loss = int(line[4])
	    else:
		c_win = 0
		c_loss = 0
	

	    data[team] = [wins,loss,perct,champ,conf,c_win,c_loss]
	    print team+": "+str(data[team])

	    # Add conference to list of conferences
	    if conf not in confern:
		confern[conf] = [team]
	    else:
		confern[conf].append(team)

	# Add entry for FCS schools played
	data['FCS'] = [0,12,'0.000','FCS']

	standings[year] = data
	conferences[year] = confern

    return standings


def import_rankings():

    rankings = {}
    rawdata = import_excel(RANKINGS)

    for year in rawdata.keys():
	#print year
	data = {}
	count = 0
	for line in rawdata[year]:
	    if count == 0:
		count+=1
		continue

	    # Pull Needed Data
	    rank = int(line[0])
	    team1 = str(line[1])
	    team2 = str(line[3])

	    # Build/Update Dictionary
	    if team1 == team2:
		# consensus ranking
		data[team1] = rank+rank
	    else:
		# Check if team is already stored and update/add as needed
		if team1 in data.keys():
		    data[team1] = data[team1] + rank
		else:
		    data[team1] = rank
		
		if team2 in data.keys():
		    data[team2] = data[team2] + rank
		else:
		    data[team2] =  rank

	rankings[year] = data

    return rankings


def import_data():

    print "\n==== IMPORTING DATA FILES ===="

    rank = import_rankings()
    stand = import_standings()
    stats = import_stats()

    print "==== COMPLETE ====\n"

    return rank, stand, stats


def trim_dataset(standings, stats):

    trimmed = {}

    for year in standings.keys():
	print year
	data = {}
	count = 0

	for team in standings[year]:
	    stand = standings[year][team]
	    loss = stand[1]
	    print team +" has "+str(loss)+" loss(es)"
	    if loss <= 4:
		print "   Keeping "+team+" Data"
		data[team] = stats[year][team]
		data[team]['Record'] = stand
		print data[team]
	    else:
		print "   Discarding "+team+" Data"

	trimmed[year] = data

    return trimmed

def strength_of_schedule(team, year, root=True):

    # Get Gamedata for selected team
    games = import_game_data(team, year)

    mov = 0.0
    owp = 0.0
    oowp = 0.0 
    sov = 0.0
    sol = 0.0
    count = 0
    wins = 0
    loss = 0
    for game in games:
	
	count += 1
	diff = game[2] - game[3]
	site = game[4]

	opp = game[1]
	if opp in espn_names.keys() and espn_names[opp] in standings[year].keys():
	    opp = espn_names[opp]
	else:
	    #print "opp is FCS"
	    opp = 'FCS'
	#print "Opponent: "+opp

	# Margin of Victory
	#   **losses will negatively affect mov ratio**
	mov += diff

	opp_stand = standings[year][opp]
	win_perct = float(opp_stand[2])

	# Strength of Victory/Losses
	if diff > 0:
	    sov += pow(win_perct, 2.5)
	    wins += 1

	    if win_perct > .7 and site == 'V':		# Quality Win on road
		wins -= 1
	else:
	    sol += pow(win_perct, 2.5)
	    loss += 1

	    if win_perct < .5 and site == 'H':		# Bad loss at home
		loss += 1

	# Strength of Schedule
	#print "opp win%: "+str(win_perct)
	#print "weighted win%: "+str(pow(win_perct, 2.5))
	owp += pow(win_perct, 2.5)

	if root:
	    if opp in cfb_names.keys():
		opponent = cfb_names[opp]

		if opp in standings[year].keys():
		    oowp += strength_of_schedule(opponent, year, False)
		else:
		    print "TEAM IS FCS AT TIME OF GAME"

	    else:
		#print "opp is FCS: no need to call SOS"
		opp = 'FCS'

		#print standings[year][opp][2]
		#print "\n"
		oowp += float(standings[str(year)][opp][2])

    # Calculate Averages
    avg_mov = float(mov) / count
    avg_owp = float(owp) / count

    # Strength of Victory		# take into account 0 wins
    if wins == 0:
	avg_sov = 0.0
    else:    
	avg_sov = (float(sov) / wins) * 100

    # Strength of Losses		# take into account 0 losses
    if loss == 0:
	avg_sol = 100.0
    else:
	avg_sol = (float(sol) / loss) * 100

    #print "\navg mov: "+str(avg_mov)
    #print "avg owp: "+str(avg_owp)
    #print "root is "+str(root)

    if root:
	#print "\n\nCalculating sos"
	avg_oowp = float(oowp) / count
	#print "avg owp: "+str(avg_owp)
	#print "avg oowp: "+str(avg_oowp)

	sos = (((2*avg_owp) + avg_oowp) / 3) * 100

	#print team+" STRENGTH OF SCHEDULE for "+str(year)+": "+str(sos)+"\n\n"
	return sos, avg_mov, avg_sov, avg_sol

    else:
	return avg_owp

def strength_of_conference(year, conf):

    ooc_w = 0
    ooc_g = 0
    for team in conferences[year][conf]:
	
	data = standings[year][team]
	wins = data[0]
	loss = data[1]
	c_wins = data[5]
	c_loss = data[6]

	ooc_w += wins - c_wins
	ooc_g += (wins+loss) - (c_wins+c_loss)

    c_win_perct = float(ooc_w) / ooc_g

    #print "ooc_wins: "+str(ooc_w)
    #print "ooc_games: "+str(ooc_g)
    #print conf+" win%: "+str(c_win_perct)

    return c_win_perct

def import_game_data(team, year):

    filename = 'Data/cfb'+str(year)+'stats.csv'

    #print team
    #print filename

    if team in cfb_names.keys():
	team = cfb_names[team]

    games = []
    count = -1

    with open(filename, 'rU') as infile:
	reader = csv.reader(infile)
	
	for line in reader:
	    if count==-1:
		count+=1
		continue

	    if team == str(line[1]):

		count += 1

		date = str(line[0])
		score_fo = int(line[2])		
		opp = str(line[10])
		score_ag = int(line[11])
		site = str(line[19])

		games.append([date,opp,score_fo,score_ag,site])

	    else:
		continue

    return games


def ranks():

    metrics = {}
   
    for year in rankings:
	print year
	data = {}
	for team in rankings[year]:
	    print team
	    sos, mov, sov, sol = strength_of_schedule(team,year)
	    votes = rankings[year][team]
	    print "-----"
	    print "Votes: "+str(votes)
	    print "SOS: "+str(sos)
	    print "SOV: "+str(sov)
	    print "SOL: "+str(sol)
	    print "MOV: "+str(mov)+"\n"

	    data[team] = [votes,sos,sov,sol,mov]
	    
    	#data = sorted(data, key=lambda team: team[0])
	metrics[year] = data 
	    
    return metrics

def observe_stats():

    file = "top25_metrics.csv"
    out = open(file, 'w')

    header = "year,team,votes,wins,loss,champ,sos,sov,sol,mov,"
    header += "off_rank,off_points,off_pts/g,"
    header += "def_rank,def_points,def_pts/g,"
    header += "pass_rank,rush_rank\n"
    out.write(header)

    # Get Data
    data = ranks()

    for year in data:
	for team in data[year]:

	    votes = str(data[year][team][0])
	    sos = str(data[year][team][1])
	    sov = str(data[year][team][2])
	    sol = str(data[year][team][3])
	    mov = str(data[year][team][4])

	    # Standings
	    wins = str(standings[year][team][0])
	    loss  = str(standings[year][team][1])
	    champ  = str(standings[year][team][3])

	    # Offense
	    off_rk = str(stats[year][team]['Offense'][0])
	    pts_for  = str(stats[year][team]['Offense'][7])
	    pts_pg  = str(stats[year][team]['Offense'][8])
	    rush_rk = str(stats[year][team]['Rushing'][0])
	    pass_rk = str(stats[year][team]['Passing'][0])

	    # Defense
	    def_rk = str(stats[year][team]['Defense'][0])
	    pts_ag  = str(stats[year][team]['Defense'][7])
	    pts_apg  = str(stats[year][team]['Defense'][8])

	    text = year+","+team+","+votes+","+wins+","+loss+","+champ+","
	    text += sos+","+sov+","+sol+","+mov+","
	    text += off_rk+","+pts_for+","+pts_pg+","
	    text += def_rk+","+pts_ag+","+pts_apg+","
	    text += pass_rk+","+rush_rk+"\n"

	    out.write(text)

	out.write("\n")

    out.close()

def ranking_algorithm(dataset):

    team_ranks = []
    year = TEST_YEAR

    print "\n\n  INDIVIDUAL TEAM ANALYSIS"
    print "==============================\n"

    for team in dataset.keys():

	print team
	print "---------------------"

	sos, mov, sov, sol = strength_of_schedule(team,year)
	wins = standings[year][team][0]
	loss = standings[year][team][1]
	conf = standings[year][team][4]
	conf_champ = standings[year][team][3]
	pts_pg  = stats[year][team]['Offense'][8]
	pts_apg  = stats[year][team]['Defense'][8]

	# Conference Champion Multiplier
	soc = strength_of_conference(year, conf)
	champ = soc*1.5
	if conf_champ == 'Y':		# Champion
	    print "Conference Champ"
	    champ += 5 * soc
	elif conf_champ == 'C':		# Co-champion
	    print "Co-conference Champ"
	    champ += 3 * soc

	PR = 1.75*(pts_pg / pts_apg) 				# Points Ratio w/Multiplier
	SMV = ((sos*mov)/100) 					# Strength of Margin of Victory
	SR = (((sov/50) * wins) - (pow(sol/100, -1) * loss))	#Strength of Record
	CF = champ						# Conference Factor 

	score = PR + SMV + SR + CF

	print "Score: "+str(score)
	print "   Points Ratio: "+str(PR)
	print "   Strength of Margin of Victory: "+str(SMV)
	print "   Strength of Record: "+str(SR)
	print "   Conf. Factor: "+str(CF)
	print "Margin of Victory: "+str(mov)
	print "Strength of Schedule: "+str(sos)
	print "Strength of Conference: "+str(soc)
	print "Strength of Victory: "+str(sov)
	print "Strength of Losses: "+str(sol)
	print "Points Scored / game: "+str(pts_pg)
	print "Points Allowed / game: "+str(pts_apg)
	print "Losses: "+str(loss)
	print "\n"

	entry = (team, score)
	team_ranks.append(entry)

    temp = sorted(team_ranks, key=lambda team: team[1], reverse=True)

    print "        FINAL RANKINGS        "
    print "=============================="
    pp.pprint(temp)


def main(argv):

    global rankings
    global standings
    global stats
    global namemap
    global espn_names
    global cfb_names


    rankings, standings, stats = import_data()
    dataset = trim_dataset(standings, stats)
    espn_names, cfb_names = team_mapping()

    # Print Statistical data to a .csv file
    #observe_stats()

    ranking_algorithm(dataset[TEST_YEAR])



if __name__ == "__main__":
    main(sys.argv[1:])
