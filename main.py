from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment
import datetime
#import json
#from pprint import pprint



NAME = 0
KLASS = 1
KLUBB = 2

dt = 3
workbook_sheets = ["Herr", "Dam", "Herr U23", "Dam U23"]
workbooks_created = []
final_results_workbook_name = "Syratomten Total Poängställning.xlsx"


"""
Functions
"""

def scoreboard(namn, klass, klubb, tid, number_of_participants, position):

    points = 5 + number_of_participants - position

    if tid != None: # If the tid is None, they have not raced

        score_workbook[klass]["A" + str(position+2)] = position + 1
        score_workbook[klass]["A" + str(position+2)].alignment = Alignment(horizontal='left')

        score_workbook[klass]["B" + str(position+2)] = namn # Name

        if klubb != None: # Dont write "None" as the club
            score_workbook[klass]["C" + str(position+2)] = klubb # Klubb


        score_workbook[klass]["D" + str(position+2)] = tid  # Tid
        score_workbook[klass]["E" + str(position+2)] = "{:.1f}".format(19.5/(int(datetime.datetime.strptime(tid, "%M:%S").strftime("%M"))/60 + int(datetime.datetime.strptime(tid, "%M:%S").strftime("%S"))/3600)) # speed


        if klubb == "Väsby SS Triathlon": # Only Väsby Triathlon members gets a score
            score_workbook[klass]["F" + str(position+2)] = points      # points
            score_workbook[klass]["F" + str(position+2)].alignment = Alignment(horizontal='left')

        return position + 1
    else: # If tid was none, dont increase the position
        return position

def fill_final_results():

    # Load the final score workbook
    final_workbook = load_workbook(filename="Syratomten Total Poängställning.xlsx")

    final_result_dict = {}

    for races in workbooks_created:

        # Load the individual race score workbook
        race_workbook = load_workbook(filename=races)
        print("INFO: Opened workbook " + races)

        # For each class in the race
        for race_class in workbook_sheets:

            # Append the workbook sheet values to a list because it is easier to work with
            for values in race_workbook[race_class].iter_rows(min_row=2, values_only=True):

                if race_class not in final_result_dict: # If the class doesn't exist in the dictionary, create the class
                    final_result_dict[race_class] = {}

                if values[5] != None: # Dont save the result if they dont have a score
                    if values[1] not in final_result_dict[race_class]: # if the participant doesn't exists
                        final_result_dict[race_class][values[1]] = {races: values[5]}
                    else:
                        final_result_dict[race_class][values[1]][races] = values[5]

    # This dictionary is translating the names of the races to a specific column in the final wourkbook.
    column_dict = {
        "Syratomten Deltävling 1.xlsx" : "C",
        "Syratomten Deltävling 2.xlsx" : "D",
        "Syratomten Deltävling 3.xlsx" : "E",
        "Syratomten Deltävling 4.xlsx" : "F",
        "Syratomten Deltävling 5.xlsx" : "G",
        "Syratomten Deltävling 6.xlsx" : "H",
        "Syratomten Deltävling 7.xlsx" : "I",
        "Syratomten Deltävling 8.xlsx" : "J",
        "Syratomten Deltävling 9.xlsx" : "K",
        "Syratomten Deltävling Final.xlsx" : "L"
    }

    for klass in final_result_dict:
        final_position = 2
        for namn in final_result_dict[klass]:
            total_points = 0
            for race in final_result_dict[klass][namn]:
                column = column_dict[race]
                final_workbook[klass]["A" + str(final_position)] = final_position-1
                final_workbook[klass]["A" + str(final_position)].alignment = Alignment(horizontal='left')

                final_workbook[klass]["B" + str(final_position)] = namn

                final_workbook[klass][column + str(final_position)] = final_result_dict[klass][namn][race]
                total_points += final_result_dict[klass][namn][race]

                final_workbook[klass][column + str(final_position)].alignment = Alignment(horizontal='center')

            final_workbook[klass]["M" + str(final_position)] = total_points
            final_workbook[klass]["M" + str(final_position)].alignment = Alignment(horizontal='left')
            final_position += 1

    # Save the final_workbook after all the results are saved
    final_workbook.save(filename="Syratomten Total Poängställning.xlsx")
    print("INFO: The workbook Syratomten Total Poängställning.xlsx was saved.")





def sort_fuction(elem):
    if elem[dt]: # If the values is not None, return that value
        return elem[dt]
    else: # If the value is None return "00:00" instead. Otherwise the sort() function will try to sort None, which doesn't work
        return "00:00"

def sort_final_score(elem):
    """if elem[12]: # If the values is not None, return that value
        return elem[12]
    else: # If the value is None return "00:00" instead. Otherwise the sort() function will try to sort None, which doesn't work
        return "00:00"""
    return elem[12]

def create_race_workbook(workbook_name):

    # Create a workbook for each race
    workbook = Workbook()
    workbook.save(filename="Syratomten " + workbook_name + ".xlsx")
    print("INFO: Workbook Syratomten " + workbook_name + ".xlsx was created.")

    # Save the name of the workbooks created so I can open them later for the final results
    workbooks_created.append("Syratomten " + workbook_name + ".xlsx")

    score_workbook = load_workbook(filename="Syratomten " + workbook_name + ".xlsx")

    # Populate the workbook with the sheets listed in the list "workbook_sheets"
    for sheet in workbook_sheets:
        if sheet not in workbook.sheetnames:
            workbook.create_sheet(sheet)

            workbook[sheet]["A1"] = "Placering"
            workbook[sheet]["B1"] = "Namn"
            workbook[sheet]["C1"] = "Klubb"
            workbook[sheet]["D1"] = "Tid"
            workbook[sheet]["E1"] = "Hastighet (km/h)"
            workbook[sheet]["F1"] = "Poäng"

            # Setting column width in the workbook sheets
            workbook[sheet].column_dimensions["A"].width = 9
            workbook[sheet].column_dimensions["B"].width = 20
            workbook[sheet].column_dimensions["C"].width = 20
            workbook[sheet].column_dimensions["D"].width = 9
            workbook[sheet].column_dimensions["E"].width = 16
            workbook[sheet].column_dimensions["F"].width = 9

            print("INFO: Sheet " + sheet + " created in workbook " + str(workbook_name))

    # Remove the sheet named "Sheet", which is created by default.
    if "Sheet" in score_workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    return workbook

def create_final_results_workbook():

    # Create the workbook
    workbook = Workbook()

    workbook.save(filename=final_results_workbook_name)
    print("INFO: Workbook " + final_results_workbook_name + " was created.")

    score_workbook = load_workbook(filename=final_results_workbook_name)

    # Populate the workbook with the sheets listed in the list "workbook_sheets"
    for sheet in workbook_sheets:
        if sheet not in workbook.sheetnames:
            workbook.create_sheet(sheet)

            workbook[sheet]["A1"] = "Placering"
            workbook[sheet]["B1"] = "Namn"
            workbook[sheet]["C1"] = "DT1"
            workbook[sheet]["D1"] = "DT2"
            workbook[sheet]["E1"] = "DT3"
            workbook[sheet]["F1"] = "DT4"
            workbook[sheet]["G1"] = "DT5"
            workbook[sheet]["H1"] = "DT6"
            workbook[sheet]["I1"] = "DT7"
            workbook[sheet]["J1"] = "DT8"
            workbook[sheet]["K1"] = "DT9"
            workbook[sheet]["L1"] = "F"
            workbook[sheet]["M1"] = "Totalt"

            workbook[sheet].column_dimensions["A"].width = 9
            workbook[sheet].column_dimensions["B"].width = 20
            workbook[sheet].column_dimensions["C"].width = 4
            workbook[sheet].column_dimensions["D"].width = 4
            workbook[sheet].column_dimensions["E"].width = 4
            workbook[sheet].column_dimensions["F"].width = 4
            workbook[sheet].column_dimensions["G"].width = 4
            workbook[sheet].column_dimensions["H"].width = 4
            workbook[sheet].column_dimensions["I"].width = 4
            workbook[sheet].column_dimensions["J"].width = 4
            workbook[sheet].column_dimensions["K"].width = 4
            workbook[sheet].column_dimensions["L"].width = 4
            workbook[sheet].column_dimensions["M"].width = 6

            print("INFO: Sheet " + sheet + " created in " + final_results_workbook_name)

    # Remove the sheet named "Sheet", which is created by default.
    if "Sheet" in score_workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    workbook.save(filename=final_results_workbook_name)
    print ("INFO: Workbook " + final_results_workbook_name + " was saved")


if __name__ == "__main__":

    # Load the workbook that includes all the race results
    init_workbook = load_workbook(filename="st-test2.xlsx")

    # Append all the values in the initial workbook to a list. It is easyier to work with
    result_list = []

    for values in init_workbook["Syra Tomten"].iter_rows(min_row=2, values_only=True):
        result_list.append(values)

    # For each race in the workbook
    for race in init_workbook["Syra Tomten"].iter_rows(min_row=1, max_row=1, min_col=4, values_only=True):

        for workbooks in race:

            # Sort the result_list based on the times
            try:
                result_list.sort(key=sort_fuction)
                #print("SORTED " + str(workbooks))
            except TypeError:
                print("WARNING: DID NOT SORT " + str(workbooks))



            number_of_participants_herr = 0
            number_of_participants_dam = 0
            number_of_participants_herru23 = 0
            number_of_participants_damu23 = 0

            """
            Count the number of participants in each class in this race.
            The number of participant is used for setting the score
            """
            for participant in result_list:
                if participant[dt] != None: # Only if the time is not None, could probably match anything
                    if participant[KLASS] == "Herr" or participant[KLASS] == "Herr U23":
                        number_of_participants_herr = number_of_participants_herr + 1
                    elif participant[KLASS] == "Dam" or participant[KLASS] == "Dam U23":
                        number_of_participants_dam = number_of_participants_dam + 1

                    if participant[KLASS] == "Herr U23":
                        number_of_participants_herru23 = number_of_participants_herru23 + 1
                    elif participant[KLASS] == "Dam U23":
                        number_of_participants_damu23 = number_of_participants_damu23 + 1



            # Reset the participans position
            position_herr = 0
            position_dam = 0
            position_herru23 = 0
            position_damu23 = 0


            # Only continue if there are any participants in the race
            if number_of_participants_herr != 0 or number_of_participants_dam != 0 or number_of_participants_herru23 != 0 or number_of_participants_damu23 != 0:
            #if any(number_of_participants_herr, number_of_participants_dam, number_of_participants_herru23, number_of_participants_damu23) != 0:

                # Create a new workbook for this race
                score_workbook = create_race_workbook(workbooks)

                for stuff in result_list:

                    # Herr U23 is also counted in the Herr class
                    if stuff[KLASS] == "Herr" or stuff[KLASS] == "Herr U23":
                        position_herr = scoreboard(stuff[NAME], "Herr", stuff[KLUBB], stuff[dt], number_of_participants_herr, position_herr)

                    # Dam U23 is also counted in the Dam class
                    elif stuff[KLASS] == "Dam" or stuff[KLASS] == "Dam U23":
                        position_dam = scoreboard(stuff[NAME], "Dam", stuff[KLUBB], stuff[dt], number_of_participants_dam, position_dam)

                    if stuff[KLASS] == "Herr U23":
                        position_herru23 = scoreboard(stuff[NAME], stuff[KLASS], stuff[KLUBB], stuff[dt], number_of_participants_herru23, position_herru23)

                    elif stuff[KLASS] == "Dam U23":
                        position_damu23 = scoreboard(stuff[NAME], stuff[KLASS], stuff[KLUBB], stuff[dt], number_of_participants_damu23, position_damu23)

                # Increase the deltävling by one each loop
                dt += 1

                # Save the content in the score workbook
                score_workbook.save(filename="Syratomten " + workbooks + ".xlsx")
                print ("INFO: Workbook Syratomten " + workbooks + ".xlsx was saved")

    # Create the final results workbook
    create_final_results_workbook()

    # Fill the final resulults workbook
    fill_final_results()

    """
    Testing some stuff
    """

    # Load the workbook that includes all the race results
    real_final_workbook = load_workbook(filename=final_results_workbook_name)
    print("INFO: Opened " + final_results_workbook_name + " to sort it.")



    for sheet in workbook_sheets:
        #print(real_final_workbook[race_class])

        # Append all the values in the initial workbook to a list. It is easyier to work with
        real_result_list = []

        for values in real_final_workbook[sheet].iter_rows(min_row=2, values_only=True):
            real_result_list.append(values)

        # For each race in the workbook
        for race in real_final_workbook[sheet].iter_rows(min_row=1, max_row=1, min_col=4, values_only=True):

            for workbooks in race:

                # Sort the result_list based on the times
                try:
                    real_result_list.sort(key=sort_final_score, reverse = True)

                except TypeError:
                    print("WARNING: DID NOT SORT " + str(workbooks))

        # Sorted list
        #print(real_result_list)


        column_dict1 = {
            0 : "A",
            1 : "B",
            2 : "C",
            3 : "D",
            4 : "E",
            5 : "F",
            6 : "G",
            7 : "H",
            8 : "I",
            9 : "J",
            10 : "K",
            11 : "L",
            12 : "M"
        }


        for position,stuff in enumerate(real_result_list,1): #enumerate starts at 1
            for idx, items in enumerate(stuff):
                column = column_dict1[idx]
                real_final_workbook[sheet][str(column) + str(position+1)] = items
                real_final_workbook[sheet]["A" + str(position+1)] = position

    real_final_workbook.save(filename=final_results_workbook_name)
    print("INFO: The workbook Syratomten Total Poängställning.xlsx was saved.")
