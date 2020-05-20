"""
Participants needs unique names
"""

DEBUG = True


import datetime
from time import strftime
from time import gmtime
import sys
import json
import copy

from module.googleapi import Google

race_name = "Simtomten"
race_headings_list = [["Placering", "Namn", "Klubb" ,"Tid", "Hastighet (tid/100m)", "Poäng"]]
final_headings_list = [["Placering", "Namn", "DT1", "DT2", "DT3", "DT4", "DT5", "F", "Totalt"]]
races_list = ['Deltävling 1', 'Deltävling 2', 'Deltävling 3', 'Deltävling 4', 'Deltävling 5', 'Final']
classes = ["Herr", "Dam"]
RACE_LENGTH = 1500 # 1500 meters
column_width_list = [80,220,220,120,150,100]


def new_sort_individual_race(elem):
    return elem[2]


def new_sort_final_score(elem):
    return elem[len(elem)-1]


def calculate_speed(time):
    min2sec = int(datetime.datetime.strptime(time, "%M:%S").strftime("%M"))*60 # Converts minutes to seconds
    sec = int(datetime.datetime.strptime(time, "%M:%S").strftime("%S"))

    number_of_hundreds = RACE_LENGTH / 100

    speed_in_seconds = (min2sec + sec) / number_of_hundreds

    speed = strftime("%M:%S", gmtime(speed_in_seconds)) # Converts seconds to minutes and seconds

    if speed[0] == "0": # exclude the first 0 in the speed, if there is one.
        return speed[1:]

    return speed

def calulcate_score(participants, position):
    score =  5 + participants - position
    return score



def make_list_certian_length(list, length):
    """
    Makes a list within a list a certain length.
    """
    for column in list:
        while len(column) <= length:
          column.append("")

def adjust_column_width(race_spreadsheet, column_width_list):
    
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
            "spreadsheetId": "1qLKWsnyhOSSAnsZ0lY0NLyen1m53kLOd_RDWa26OqLI",
            "range": "!A2:M",
            "sheetName": "Test"
        }
    elif spreadsheet_var == "2020":

        google_sheet = {
            "spreadsheetId": "1qLKWsnyhOSSAnsZ0lY0NLyen1m53kLOd_RDWa26OqLI",
            "range": "!A2:M",
            "sheetName": "2020"
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
            update_user_list.append(speed)

            if participant["class"] == "Herr":
                herr_race_list.append(update_user_list)
                herr_number_of_participants += 1
            elif participant["class"] == "Dam":
                dam_race_list.append(update_user_list)
                dam_number_of_participants += 1


        if DEBUG: print(f"Number of participants in {race['race']}: Herr: {herr_number_of_participants}, Dam: {dam_number_of_participants}")

        if herr_number_of_participants == 0 and dam_number_of_participants == 0:
            continue # If there are not participants in the race, continue to the next race.

        race_spreadsheet = Google.create_spreadsheet('Poängräkning ' + race_name + ' ' + race["race"] + ' ' + spreadsheet_var, sheet_titles_list)
        race["spreadsheet_id"] = race_spreadsheet
        #print(race_spreadsheet)
        print(f'Created spreadsheet for {race["spreadsheet_id"]}')

        
        column_width_data = adjust_column_width(race["spreadsheet_id"], column_width_list)
        #print(column_width_data)
        
        Google.batch_update(race["spreadsheet_id"], column_width_data)
        print(f"Spreadsheet {race['race']} was updated with correct column width.")


        


        # Sort the lists based on the result
        herr_race_list.sort(key=new_sort_individual_race)
        dam_race_list.sort(key=new_sort_individual_race)

        position_herr = 0
        position_dam = 0

        for participant in herr_race_list:
            participant.insert(0, str(position_herr+1)) # Adding the position as string so the text will automatically be left aligned.
            if "Väsby SS Triathlon" in participant: # Lazy search for Väsby SS Traithlon as the club. Only members of Väsby SS Traithlon get points
                score = calulcate_score(herr_number_of_participants, position_herr)
                participant.append(str(score)) # Adding the score as string so the text will automatically be left aligned.
            position_herr +=1

        for participant in dam_race_list:
            participant.insert(0, str(position_dam+1)) # Adding the position as string so the text will automatically be left aligned.
            if "Väsby SS Triathlon" in participant: # Lazy search for Väsby SS Traithlon as the club. Only members of Väsby SS Traithlon get points
                score = calulcate_score(dam_number_of_participants, position_dam)
                participant.append(str(score)) # Adding the score as string so the text will automatically be left aligned.
            position_dam +=1


        # Update the race spreadsheets with the results
        for race_class in classes:
            google_data = []
            if race_class == "Herr":
                google_data = race_headings_list + herr_race_list
                Google.update(race_spreadsheet, race_class, "!A1:M", google_data)
                print(f"Updated spreadsheet {race['race']} ({race_spreadsheet}) for class {race_class} with the score")
            elif race_class == "Dam":
                google_data = race_headings_list + dam_race_list
                Google.update(race_spreadsheet, race_class, "!A1:M", google_data)
                print(f"Updated spreadsheet {race['race']} ({race_spreadsheet}) for class {race_class} with the score")

    print("Opening each race spreadsheet and saving all the results in a dictionary.")

    final_result_dict = {}
    # Open each race score spreadsheet for each race and read the score of each participant
    for race in total_race_result_list: # For each race. (This just extracts the spreadsheet_id)
        #print(f'Opening spreadhseet for {race["race"]} and saving it')
        for race_class in classes: # For each race_class within the race
            if "spreadsheet_id" in race: # spreadsheet_id only exists if there is a sheet created for this race.
                individual_race_result = Google.get(spreadsheet_id=race["spreadsheet_id"], sheet_name=race_class)
                print(f'Opened spreadsheet {race["spreadsheet_id"]} for {race["race"]} {race_class  }')

                # The score is missing for the participant if it is a member of Väsby SS Triathlon
                make_list_certian_length(individual_race_result, 5)

                # participant[1] = namn of participant
                # participant[5] = participant's score
                for participant in individual_race_result:

                    if race_class not in final_result_dict: # If the class doesn't exist in the dictionary, create the class
                        final_result_dict[race_class] = {}
                        if DEBUG: print(f"Class {race_class} created in final_result_dict")

                    if participant[5] != "": # Dont save the result if they dont have a score
                        #print(f'{participant[1]} : {race["race"]} : {participant[5]}')
                        if participant[1] not in final_result_dict[race_class]: # if the participant doesn't exists in the dictionary already
                            final_result_dict[race_class][participant[1]] = {race["race"]: participant[5]}
                        else:
                            final_result_dict[race_class][participant[1]][race["race"]] = participant[5]
    
    
    final_spreadsheet_id = Google.create_spreadsheet(race_name + " Total Poängställning", sheet_titles_list)
    print(f"Created spreadsheet " + race_name + " Total Poängställning")
    
    # Updating the final score spreadsheet with the headings.
    for race_class in final_result_dict:
        Google.update(final_spreadsheet_id, race_class, "!A1:M", final_headings_list)
        print(f"Updated spreadsheet " + race_name + " Total Poängställning for class {race_class} with the headings")
    

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
        print(f"Updated " + race_name + " Total Poängställning for class {race_class} with the score information.")