"""
Participants needs unique names
"""

import datetime
import sys
import json
import copy
import time

from module.googleapi import Google

race_headings_list = [["Placering", "Namn", "Klubb" ,"Tid", "Hastighet (km/h)", "Poäng"]]
final_headings_list = [["Placering", "Namn", "DT1", "DT2", "DT3", "DT4", "DT5", "DT6", "DT7", "DT8", "DT9", "F", "Totalt"]]
races_list = ['Deltävling 1', 'Deltävling 2', 'Deltävling 3', 'Deltävling 4', 'Deltävling 5', 'Deltävling 6', 'Deltävling 7', 'Deltävling 8', 'Deltävling 9', 'Final']
classes = ["Herr", "Dam"]
RACE_LENGTH = 19.53 # race length in kilometers
race_column_width_list = [80,220,220,120,150,100]
final_column_width_list = [83,200,40,40,40,40,40,40,40,40,40,40,53]


def new_sort_individual_race(elem):
    return elem[2]


def new_sort_final_score(elem):
    return elem[11]


def calculate_speed(time): # Time needs to be in HH:MM format
    min2hrs = int(datetime.datetime.strptime(time, "%M:%S").strftime("%M"))/60
    sec2hrs = int(datetime.datetime.strptime(time, "%M:%S").strftime("%S"))/3600
    speed = RACE_LENGTH / (min2hrs + sec2hrs)
    return speed

def calculcate_score(participants, position):
    return 5 + participants - position

def calculcate_score2(participants, position):

    if position+1 == 1:
        extra = 5
    elif position+1 == 2:
        extra = 3
    elif position+1 == 3:
        extra = 1
    else:
        extra = 0
    return 5 + participants - position + extra

def make_list_certian_length(list, length):
    """
    Makes a list within a list a certain length.
    """
    for column in list:
        while len(column) <= length:
          column.append("")

"""def make_list_certian_length2(the_list, length):
    
    #Makes a list a certian length.
    
    while len(the_list) <= length:
        the_list.append("")"""

def adjust_column_width(race_spreadsheet, column_width_list):
    """
    Takes the spreadsheet_id and a list with all column width
    and prepares the request for updating the column width in a sheet.
    """
    
    sheet_id_list = Google.get_sheetid(race_spreadsheet)
    
    requests= {}
    requests["requests"] = []

    for sheet_id in sheet_id_list:
        for column_idx, column_width in enumerate(column_width_list):
            temp_google_data = {}
            temp_google_data = {
                "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": column_idx,
                    "endIndex": column_idx+1
                },
                "properties": {
                    "pixelSize": column_width
                },
                "fields": "pixelSize"
                }
            }
            requests["requests"].append(temp_google_data)

    return requests
        

if __name__ == "__main__":

    if len(sys.argv) == 2:
        spreadsheet_var = sys.argv[1].lower()
    else:
        print("Invalid number of argument, Enter test or the year")
        print("Example: python " + sys.argv[0].lower() + " test")
        exit(1)

    if spreadsheet_var == "test":

        google_sheet = {
            "spreadsheetId": "1a4_U99Dnk3i1HxMltCJXqkVPRabUnz_RI_85O5GYxL8",
            "range": "!A2:M",
            "sheetName": "Test"
        }
    elif spreadsheet_var == "2020":

        google_sheet = {
            "spreadsheetId": "1a4_U99Dnk3i1HxMltCJXqkVPRabUnz_RI_85O5GYxL8",
            "range": "!A2:M",
            "sheetName": "2020"
        }
    elif spreadsheet_var == "2021":

        google_sheet = {
            "spreadsheetId": "1a4_U99Dnk3i1HxMltCJXqkVPRabUnz_RI_85O5GYxL8",
            "range": "!A2:M",
            "sheetName": "2021"
        }
    elif spreadsheet_var == "2022":

        google_sheet = {
            "spreadsheetId": "1a4_U99Dnk3i1HxMltCJXqkVPRabUnz_RI_85O5GYxL8",
            "range": "!A2:M",
            "sheetName": "2022"
        }
    elif spreadsheet_var == "2023":

        google_sheet = {
            "spreadsheetId": "1a4_U99Dnk3i1HxMltCJXqkVPRabUnz_RI_85O5GYxL8",
            "range": "!A2:M",
            "sheetName": "2023"
        }
    else:
        print("No valid arguments entered. Exiting...")
        exit(1)


    main_workbook = Google.get(google_sheet["spreadsheetId"], google_sheet["sheetName"], "!A2:M")
    print("Opened the main score spreadsheet.")

    for participant in main_workbook:
        while len(participant) <= 12:
            participant.append("")

    total_race_result_list =  []
    for idx,race in enumerate(races_list):
        race_result_dict = {} # Without clearing the dictionary the data will be overwritten in the total_race_result_list
        race_result_dict["race"] = race
        race_result_dict["participants"] = []

        for participant in main_workbook:
            participant_structur = {} # Clear the dictionary
            if participant[3+idx]:
                participant_structur = {
                    "name": participant[0],
                    "class": participant[1],
                    "club": participant[2],
                    "result": participant[3+idx]
                }
                race_result_dict["participants"].append(participant_structur)


        total_race_result_list.append(race_result_dict)

    """
    To update a google spreadsheet the data needs to be a array (list within a list).
    Each user and its result has its own list. The list is cleared between each race.
    """
    sheet_titles_list = [] # The sheet titles are sent as a list in the Google API

    # Fill the sheet_titles_list with the name of the sheets that the spreadsheet should have
    for race_class in classes:
        sheet_dict = {}
        sheet_dict["properties"] = {"title" : race_class}
        sheet_titles_list.append(sheet_dict)



    for race in total_race_result_list:
        time.sleep(10)

        #race_spreadsheet = Google.create_spreadsheet(race["race"], sheet_titles_list)
        #print(f'Created spreadsheet for {race["race"]}')

        herr_race_list = []
        dam_race_list = []

        herr_number_of_participants = 0
        dam_number_of_participants = 0


        for participant in race["participants"]:
            update_user_list = []
            update_user_list.append(participant["name"])
            update_user_list.append(participant["club"])
            update_user_list.append(participant["result"])

            speed = calculate_speed(participant["result"])
            update_user_list.append("{:.1f}".format(speed))

            if participant["class"] == "Herr":
                herr_race_list.append(update_user_list)
                herr_number_of_participants += 1
            elif participant["class"] == "Dam":
                dam_race_list.append(update_user_list)
                dam_number_of_participants += 1


        print(f"Number of participants in {race['race']}: Herr: {herr_number_of_participants}, Dam: {dam_number_of_participants}")

        if herr_number_of_participants == 0 and dam_number_of_participants == 0:
            continue # If there are not participants in the race, continue to the next race.

        race_spreadsheet = Google.create_spreadsheet('Poängräkning Syratomten ' + race["race"] + ' ' + spreadsheet_var, sheet_titles_list)
        race["spreadsheet_id"] = race_spreadsheet
        
        print(f"Created spreadsheet for {race['race']}")

        # Prepare the data for changing the sheet column width
        column_width_data = adjust_column_width(race["spreadsheet_id"], race_column_width_list)
        
        # Column width is updated with a batch_update
        Google.batch_update(race["spreadsheet_id"], column_width_data)
        print(f"Updated {race['race']} with correct column width.")


        # Sort the lists based on the result
        herr_race_list.sort(key=new_sort_individual_race)
        dam_race_list.sort(key=new_sort_individual_race)

        position_herr = 0
        position_dam = 0

        for participant in herr_race_list:
            participant.insert(0, str(position_herr+1)) # Adding the position as string so the text will automatically be left aligned.
            if "Väsby SS Triathlon" in participant: # Lazy search for Väsby SS Traithlon as the club. Only members of Väsby SS Traithlon get points
                score = calculcate_score(herr_number_of_participants, position_herr)
                participant.append(str(score)) # Adding the score as string so the text will automatically be left aligned.
            position_herr +=1

        for participant in dam_race_list:
            participant.insert(0, str(position_dam+1)) # Adding the position as string so the text will automatically be left aligned.
            if "Väsby SS Triathlon" in participant: # Lazy search for Väsby SS Traithlon as the club. Only members of Väsby SS Traithlon get points
                score = calculcate_score(dam_number_of_participants, position_dam)
                participant.append(str(score)) # Adding the score as string so the text will automatically be left aligned.
            position_dam +=1


        # Update the race spreadsheets with the results
        for race_class in classes:
            google_data = []
            if race_class == "Herr":
                google_data = race_headings_list + herr_race_list
                Google.update(race_spreadsheet, race_class, "!A1:M", google_data)
                print(f"Updated spreadsheet {race['race']} for class {race_class} with the score")
            elif race_class == "Dam":
                google_data = race_headings_list + dam_race_list
                Google.update(race_spreadsheet, race_class, "!A1:M", google_data)
                print(f"Updated spreadsheet {race['race']} for class {race_class} with the score")

    print("Opening each race spreadsheet and saving all the results in a dictionary.")

    final_result_dict = {}
    # Open each race score spreadsheet for each race and read the score of each participant
    for race in total_race_result_list: # For each race. (This just extracts the spreadsheet_id)
        
        for race_class in classes: # For each race_class within the race

            if "spreadsheet_id" in race: # spreadsheet_id only exists if there is a sheet created for this race.
                individual_race_result = Google.get(spreadsheet_id=race["spreadsheet_id"], sheet_name=race_class)
                print(f'Opened spreadsheet {race["spreadsheet_id"]} for {race["race"]} {race_class}')

                # The score is missing for the participant if it is a member of Väsby SS Triathlon
                make_list_certian_length(individual_race_result, 5)

                # participant[1] = name of participant
                # participant[5] = participant's score
                for participant in individual_race_result:

                    if race_class not in final_result_dict: # If the class doesn't exist in the dictionary, create the class
                        final_result_dict[race_class] = {}

                    if participant[5] != "": # Dont save the result if they dont have a score
                        #print(f'{participant[1]} : {race["race"]} : {participant[5]}')
                        if participant[1] not in final_result_dict[race_class]: # if the participant doesn't exists in the dictionary already
                            final_result_dict[race_class][participant[1]] = {race["race"]: participant[5]}
                        else:
                            final_result_dict[race_class][participant[1]][race["race"]] = participant[5]
    
    
    final_spreadsheet_id = Google.create_spreadsheet("Syratomten Total Poängställning", sheet_titles_list)
    print(f"Created spreadsheet Syratomten Total Poängställning")
    
    # Updating the final score spreadsheet with the headings.
    for race_class in final_result_dict:
        Google.update(final_spreadsheet_id, race_class, "!A1:M", final_headings_list)
        print(f"Updated spreadsheet Syratomten Total Poängställning for class {race_class} with the headings")

    # Prepare the data for changing the sheet column width
    column_width_data = adjust_column_width(final_spreadsheet_id, final_column_width_list)
    
    # Column width is updated with a batch_update
    Google.batch_update(final_spreadsheet_id, column_width_data)
    print(f"Updated Syratomten Total Poängställning with correct column width.")
    

    """
    This for loop opens the final_result_dict where all the participants score's are stored and creates a nested list.
    The Google API is expecting a nested list. Each nested list is its own row within the spreadsheet.
    The nested list must be formatted as followed: 
    ["participants name", "score 1", "score 2", "score 3", "score 4", "score 5", "score 6", "score 7", "score 8", "score 9", "score 10", "total score"]
    """
    for race_class in final_result_dict:

        final_score_list = []

        for participant in final_result_dict[race_class]:
            participant_total_score = 0
            participant_score_list = []
            for race_name in races_list:
                if race_name in final_result_dict[race_class][participant]:
                    #print(participant, race_name, final_result_dict[race_class][participant][race_name])
                    participant_score_list.append(int(final_result_dict[race_class][participant][race_name])) # Adds participant's score to the list as an integer. Otherwise it will be an string.
                else:
                    participant_score_list.append("")

            
            for points in participant_score_list:
                if isinstance(points, int): # An empty string was added to the list when the participant has not raced and we can't sum that up.
                    participant_total_score += points # Sums upp all the participant's scores

            # Adds the participants name first in the list.
            participant_score_list.insert(0,participant)

            # Fills the gap from the race to the 11th spot in order to add the particiapant's total score on the 12th
            #make_list_certian_length2(participant_score_list, 11) 
            
            # Adds the participant's total score on 12th spot
            participant_score_list.append(participant_total_score)

            final_score_list.append(participant_score_list)


        final_score_list.sort(key=new_sort_final_score, reverse = True)

        for idx, participant in enumerate(final_score_list, 1):
            participant.insert(0, idx)

        Google.update(final_spreadsheet_id, race_class, google_sheet["range"], final_score_list)
        print(f"Updated Syratomten Total Poängställning for class {race_class} with the score information.")