import gspread
from google.oauth2 import service_account
import random
import json
import statistics

# Load service account credentials from JSON file
with open('credentials.json') as f:
    credentials_data = json.load(f)

# Extract necessary fields
client_email = credentials_data['client_email']
private_key = credentials_data['private_key']

# Set up authentication
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_info(
    {
        "client_email": client_email,
        "private_key": private_key,
        "type": "service_account",
        "project_id": "ball-hockey-schedule-randomize",
        "client_id": "143469348680-6p0p3r5rj61okrb242hpjaqaoo984j4j.apps.googleusercontent.com",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/ball-hockey-schedule-randomize.iam.gserviceaccount.com"
    },
    scopes=scope
)

# Create a client instance for Google Sheets API
client = gspread.authorize(credentials)

# Open the spreadsheet
spreadsheet = client.open('WPBHL Cost Sheet')

# Get the worksheet by title
worksheet = spreadsheet.worksheet('Fall \'23 Player Tracker')

# Get all player data from the worksheet
player_data = worksheet.get_all_records()

# Separate players into Group 4 based on values 'TRUE' in Column B
group_4 = [player for player in player_data if player['Goalie'] == 'TRUE']

# Shuffle Group 4
random.shuffle(group_4)

# Separate player names into Group 1, Group 2, and Group 3 based on eligibility in Columns C and D
group_1 = [player for player in player_data if player['7:15-8:15'] == 'TRUE' and player['8:30-9:30'] == 'FALSE' and player not in group_4]
group_2 = [player for player in player_data if player['7:15-8:15'] == 'FALSE' and player['8:30-9:30'] == 'TRUE' and player not in group_4]
group_3 = [player for player in player_data if player['7:15-8:15'] == 'TRUE' and player['8:30-9:30'] == 'TRUE' and player not in group_4]

# Shuffle each group
random.shuffle(group_1)
random.shuffle(group_2)
random.shuffle(group_3)

# Assign players from Group 1 to '7:15-8:15'
time_slot_1 = group_1[:12]

# Assign players from Group 2 to '8:30-9:30'
time_slot_2 = group_2[:12]

# Randomly assign players from Group 3 to '7:15-8:15' or '8:30-9:30' until there are 14 players in each time slot
for player in group_3:
    if len(time_slot_1) < 12 and player not in time_slot_1 and player not in time_slot_2:
        time_slot_1.append(player)
    elif len(time_slot_2) < 12 and player not in time_slot_1 and player not in time_slot_2:
        time_slot_2.append(player)

# Shuffle the player assignments within each time slot
random.shuffle(time_slot_1)
random.shuffle(time_slot_2)

# Split time_slot_1 into "Home 1" and "Away 1" teams
home_1 = time_slot_1[:6]
away_1 = time_slot_1[6:]

# Split time_slot_2 into "Home 2" and "Away 2" teams
home_2 = time_slot_2[:6]
away_2 = time_slot_2[6:]

# Calculate the mean skill value for 'Home 1' and 'Away 1' teams
home_1_skills = [player['Skill'] for player in home_1]
away_1_skills = [player['Skill'] for player in away_1]
home_1_mean = statistics.mean(home_1_skills)
away_1_mean = statistics.mean(away_1_skills)

# Check and balance the mean skill values for 'Home 1' and 'Away 1' teams
while abs(home_1_mean - away_1_mean) > 1:
    if home_1_mean > away_1_mean:
        player_to_swap_home = max(home_1, key=lambda player: player['Skill'])
        player_to_swap_away = min(away_1, key=lambda player: player['Skill'])
        home_1.remove(player_to_swap_home)
        away_1.remove(player_to_swap_away)
        home_1.append(player_to_swap_away)
        away_1.append(player_to_swap_home)
    else:
        player_to_swap_home = min(home_1, key=lambda player: player['Skill'])
        player_to_swap_away = max(away_1, key=lambda player: player['Skill'])
        home_1.remove(player_to_swap_home)
        away_1.remove(player_to_swap_away)
        home_1.append(player_to_swap_away)
        away_1.append(player_to_swap_home)

    home_1_skills = [player['Skill'] for player in home_1]
    away_1_skills = [player['Skill'] for player in away_1]
    home_1_mean = statistics.mean(home_1_skills)
    away_1_mean = statistics.mean(away_1_skills)

# Calculate the mean skill value for 'Home 2' and 'Away 2' teams
home_2_skills = [player['Skill'] for player in home_2]
away_2_skills = [player['Skill'] for player in away_2]
home_2_mean = statistics.mean(home_2_skills)
away_2_mean = statistics.mean(away_2_skills)

# Check and balance the mean skill values for 'Home 2' and 'Away 2' teams
while abs(home_2_mean - away_2_mean) > 1:
    if home_2_mean > away_2_mean:
        player_to_swap_home = max(home_2, key=lambda player: player['Skill'])
        player_to_swap_away = min(away_2, key=lambda player: player['Skill'])
        home_2.remove(player_to_swap_home)
        away_2.remove(player_to_swap_away)
        home_2.append(player_to_swap_away)
        away_2.append(player_to_swap_home)
    else:
        player_to_swap_home = min(home_2, key=lambda player: player['Skill'])
        player_to_swap_away = max(away_2, key=lambda player: player['Skill'])
        home_2.remove(player_to_swap_home)
        away_2.remove(player_to_swap_away)
        home_2.append(player_to_swap_away)
        away_2.append(player_to_swap_home)

    home_2_skills = [player['Skill'] for player in home_2]
    away_2_skills = [player['Skill'] for player in away_2]
    home_2_mean = statistics.mean(home_2_skills)
    away_2_mean = statistics.mean(away_2_skills)

# Ensure 'Home 1' and 'Away 1' consist of 6 players
while len(home_1) > 6:
    player_to_remove = home_1.pop()
    group_1.append(player_to_remove)

while len(away_1) > 6:
    player_to_remove = away_1.pop()
    group_1.append(player_to_remove)

# Ensure 'Home 2' and 'Away 2' consist of 6 players
while len(home_2) > 6:
    player_to_remove = home_2.pop()
    group_2.append(player_to_remove)

while len(away_2) > 6:
    player_to_remove = away_2.pop()
    group_2.append(player_to_remove)

# Assign two players from Group 4 to '7:15-8:15 Goalies' team
time_slot_1_goalies = []
if len(group_4) >= 2:
    time_slot_1_goalies = group_4[:2]
    group_4 = group_4[2:]
else:
    print("Not enough players in Group 4 to assign to '7:15-8:15 Goalies' team.")

# Assign twoplayers from Group 4 to '8:30-9:30 Goalies' team
time_slot_2_goalies = []
if len(group_4) >= 2:
    time_slot_2_goalies = group_4[:2]
else:
    print("Not enough players in Group 4 to assign to '8:30-9:30 Goalies' team.")

# Print the player assignments with mean skill values
print("7:15-8:15")
print("Home 1:", [player['Name'] for player in home_1], "- Mean Skill:", home_1_mean)
print("Away 1:", [player['Name'] for player in away_1], "- Mean Skill:", away_1_mean)
print("'7:15-8:15 Goalies':", [player['Name'] for player in time_slot_1_goalies])
print()
print("8:30-9:30")
print("Home 2:", [player['Name'] for player in home_2], "- Mean Skill:", home_2_mean)
print("Away 2:", [player['Name'] for player in away_2], "- Mean Skill:", away_2_mean)
print("'8:30-9:30 Goalies':", [player['Name'] for player in time_slot_2_goalies])

# Write the player assignments to the 'Schedule' worksheet
schedule_worksheet = spreadsheet.worksheet('Schedule')

# Clear the existing data in the 'Schedule' worksheet
schedule_worksheet.clear()

# Define the column headers
headers = ['Time Slot', 'Home', 'Away', 'Goalies']

# Define the data rows
data = [
    ['7:15-8:15', ', '.join([player['Name'] for player in home_1]), ', '.join([player['Name'] for player in away_1]), ', '.join([player['Name'] for player in time_slot_1_goalies])],
    ['8:30-9:30', ', '.join([player['Name'] for player in home_2]), ', '.join([player['Name'] for player in away_2]), ', '.join([player['Name'] for player in time_slot_2_goalies])]
]


# Write the column headers
schedule_worksheet.append_row(headers)

# Write the data rows
schedule_worksheet.append_rows(data)