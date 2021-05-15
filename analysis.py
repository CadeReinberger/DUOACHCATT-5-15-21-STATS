import numpy as np
import pandas as pd
import openpyxl as pyxl
import os
from dataclasses import dataclass

ROUND_DIR = 'scoresheets/'
ROUND_PREFIX = 'MATCH'

NUM_ROUNDS = 7
NUM_ROOMS = 4

class round_results:
    win = 'win'
    loss = 'loss'
    tie = 'tie'

@dataclass
class team_game:
    game_result: str
    roster: set
    score: int
    category_points: int
    alphabet_points: int
    lightning_points: int    
    tossups: int 
    powers: int
    
@dataclass 
class indiv_game:
    points: int
    tossups: int
    powers: int 

def get_round_location(_room, _round):
    return ROUND_DIR + ROUND_PREFIX + str(_round) + str(_room) + '.xlsx'

def weak_strip(s):
    return s.strip() if isinstance(s, str) else s

def weak_int(s):
    try:
        return int(s)
    except:
        return 0

def parse_team_stats(_room, _round):
    #open the excel file
    round_loc = get_round_location(_room, _round)
    workbook = pyxl.load_workbook(filename = round_loc, data_only=True)
    #define all the workbooks
    cat_sheet = workbook['Category']
    alpha_sheet = workbook['Alphabet']
    light_sheet = workbook['Lightning']
    final_sheet = workbook['Final']
    indiv_sheet = workbook['Individuals']
    #get the team names
    a_name = weak_strip(cat_sheet['C1'].value)
    b_name = weak_strip(cat_sheet['F1'].value)
    #get the rosters
    a_roster = set()
    for col in ['C', 'D']:
        player = weak_strip(cat_sheet[col + '2'].value)
        if player != '' and player != 0: a_roster.add(player) 
    b_roster = set()
    for col in ['F', 'G']:
        player = weak_strip(cat_sheet[col + '2'].value)
        if player != '' and player != 0: b_roster.add(player)         
    #get the category points
    a_cat_points = weak_int(cat_sheet['E33'].value)
    b_cat_points = weak_int(cat_sheet['H33'].value)
    #get the alphabet points
    a_alpha_points = weak_int(alpha_sheet['B3'].value)
    b_alpha_points = weak_int(alpha_sheet['F3'].value)
    #get the lightning points
    a_light_points = weak_int(light_sheet['E33'].value)
    b_light_points = weak_int(light_sheet['H33'].value)
    #get the number of regulars and powers for each team
    a_tossups = 0
    a_powers = 0
    b_tossups = 0
    b_powers = 0
    for row in range(2, 6):
        r = str(row)
        name = weak_strip(indiv_sheet['A' + r].value)
        tossups = weak_int(indiv_sheet['D' + r].value)
        powers = weak_int(indiv_sheet['C' + r].value)
        if name in a_roster:
            a_tossups += tossups
            a_powers += powers
        if name in b_roster:
            b_tossups += tossups
            b_powers += powers
    #get the scores
    a_score = weak_int(final_sheet['C9'].value)
    b_score = weak_int(final_sheet['H9'].value)
    #get the game result
    if a_score > b_score: 
        a_result = round_results.win
        b_result = round_results.loss
    elif b_score > a_score:
        a_result = round_results.loss
        b_result = round_results.win
    elif a_score == 0 and b_score == 0:
        return None
    else: 
        a_result = round_results.tie
        b_result = round_results.tie
    #define the team_games
    a_game = team_game(game_result = a_result,
                       roster = a_roster, 
                       score = a_score, 
                       category_points = a_cat_points,
                       alphabet_points = a_alpha_points,
                       lightning_points = a_light_points,
                       tossups = a_tossups,
                       powers = a_powers)
    b_game = team_game(game_result = b_result,
                       roster = b_roster, 
                       score = b_score, 
                       category_points = b_cat_points,
                       alphabet_points = b_alpha_points,
                       lightning_points = b_light_points,
                       tossups = b_tossups,
                       powers = b_powers)
    #return the result
    result = {a_name : a_game, 
              b_name : b_game}
    return result

def parse_indiv_stats(_room, _round):
    #open the excel file
    round_loc = get_round_location(_room, _round)
    workbook = pyxl.load_workbook(filename = round_loc, data_only=True)
    final_sheet = workbook['Final']
    indiv_sheet = workbook['Individuals']
    #check if the game occurred
    a_score = weak_int(final_sheet['C9'].value)
    b_score = weak_int(final_sheet['H9'].value)
    if a_score == 0 and b_score == 0:
        return None
    #Read through all the players to get the data
    player_data = {}
    for row in range(2, 6):
        #read the player data
        r = str(row)
        name = weak_strip(indiv_sheet['A' + r].value)
        points = weak_int(indiv_sheet['B' + r].value)
        powers = weak_int(indiv_sheet['C' + r].value)
        tossups = weak_int(indiv_sheet['D' + r].value)
        if name == '' or name == 0:
            #sheet not all the way filled, skip
            continue
        player_game = indiv_game(points=points,
                                 tossups=tossups,
                                 powers=powers)
        player_data[name] = player_game
    return player_data   

def blank_team_internal_df():
    return {'win' : 0, 
            'loss' : 0,
            'tie' : 0,
            'points' : 0,
            'cat_points' : 0,
            'alpha_points' : 0,
            'light_points' : 0,
            'tossups' : 0,
            'powers' : 0,
            'roster' : set()}

def get_teams_dataframe():
    teams_df = {}
    for room_num in range(NUM_ROOMS):
        for round_num in range(1, NUM_ROUNDS + 1):
            round_res = parse_team_stats(room_num, round_num)
            if round_res is None:
                continue
            for (team, stats) in round_res.items():
                #if it's the team's first time, fill them with zeros
                if not team in teams_df:
                    teams_df[team] = blank_team_internal_df()
                #add the game to their record
                teams_df[team][stats.game_result] += 1
                #add to the points totals
                teams_df[team]['points'] += stats.score
                teams_df[team]['cat_points'] += stats.category_points
                teams_df[team]['alpha_points'] += stats.alphabet_points
                teams_df[team]['light_points'] += stats.lightning_points
                #add the tossups and power numbers
                teams_df[team]['tossups'] += stats.tossups
                teams_df[team]['powers'] += stats.powers
                #update the roster
                teams_df[team]['roster'] = teams_df[team]['roster'].union(stats.roster)
    return teams_df

def blank_indiv_internal_df():
    return {'points'  : 0,
            'tossups' : 0,
            'powers' : 0, 
            'games' : 0}

def get_indiv_dataframe():
    indiv_df = {}
    for room_num in range(NUM_ROOMS):
        for round_num in range(1, NUM_ROUNDS + 1):
            round_res = parse_indiv_stats(room_num, round_num)
            if round_res is None:
                continue
            for (player, stats) in round_res.items():
                #if it's the team's first time, fill them with zeros
                if not player in indiv_df:
                    indiv_df[player] = blank_indiv_internal_df()
                #add one to the games played
                indiv_df[player]['games'] += 1
                #add to the points totals
                indiv_df[player]['points'] += stats.points
                #add the tossups and power numbers
                indiv_df[player]['tossups'] += stats.tossups
                indiv_df[player]['powers'] += stats.powers
    return indiv_df

def get_roster_display_df(rosters_dict):
    rost_dict = {team : pd.Series(sorted(list(players))) for (team, players) in rosters_dict.items()}
    rost_df = pd.DataFrame(rost_dict)
    rost_df = rost_df.transpose()
    #rost_df.columns = ['' for _ in range(len(rost_df.columns))] bad
    return rost_df

def get_team_display_df():
    tddf = get_teams_dataframe()
    rosters = {}
    for (team, stats) in tddf.items():
        tddf[team]['games'] = stats['win'] + stats['loss'] + stats['tie']
        pg = lambda x : x/tddf[team]['games'] 
        tddf[team]['pct'] = pg(stats['win'] + .5 * stats['tie'])
        tddf[team]['ppg'] = pg(stats['points'])
        tddf[team]['powpg'] = pg(stats['powers'])
        tddf[team]['tupg'] = pg(stats['tossups'])
        tddf[team]['cat_ppg'] = pg(stats['cat_points'])
        tddf[team]['alpha_ppg'] = pg(stats['alpha_points'])
        tddf[team]['light_ppg'] = pg(stats['light_points'])
        rosters[team] = tddf[team]['roster']
        del(tddf[team]['roster']) #for now, we can incorproate this in a minute
    #make it a pandas dataframe, tranpose, and set the column order
    tddf = pd.DataFrame(tddf)
    tddf = tddf.T
    tddf = tddf[['win', 'loss', 'tie', 'pct', 'ppg', 'powpg', 'tupg', 
                 'cat_ppg', 'alpha_ppg', 'light_ppg', 'points', 'powers', 
                 'tossups', 'cat_points','alpha_points', 'light_points']]
    #sort the teams by winning percentage. MergeSort because it's stable
    tddf = tddf.sort_values(by='pct', axis=0, ascending=False, kind='mergesort')
    return tddf, get_roster_display_df(rosters)

def get_indiv_display_df():
    iddf = get_indiv_dataframe()
    for (player, stats) in iddf.items():
        pg = lambda x : x/stats['games'] 
        iddf[player]['ppg'] = pg(stats['points'])
        iddf[player]['powpg'] = pg(stats['powers'])
        iddf[player]['tupg'] = pg(stats['tossups'])
    #make it a pandas dataframe, tranpose, and set the column order
    iddf = pd.DataFrame(iddf)
    iddf = iddf.T
    iddf = iddf[['ppg', 'powpg', 'tupg', 'powers', 'tossups', 'points', 'games']]
    #sort the teams by winning percentage. MergeSort because it's stable
    iddf = iddf.sort_values(by='ppg', axis=0, ascending=False, kind='mergesort')
    return iddf

def combine_and_write_stats():
    tddf, rddf = get_team_display_df()
    iddf = get_indiv_display_df()
    with pd.ExcelWriter('combined_stats.xlsx') as writer:  
        tddf.to_excel(writer, sheet_name='Team Stats')
        iddf.to_excel(writer, sheet_name='Individual Stats')
        rddf.to_excel(writer, sheet_name='Team Rosters', header=False)

combine_and_write_stats()
print('Stats Successfully Compiled.')
        
    