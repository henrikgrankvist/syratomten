from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment
import datetime
import json

init_workbook = load_workbook(filename="st-test2.xlsx")

NAME = 0
KLASS = 1
KLUBB = 2

dt = 3
workbook_sheets = ["Herr", "Dam", "Herr U23", "Dam U23"]


"""
Functions
"""

def scoreboard(namn, klass, klubb, tid, number_of_participants, position):

    points = 5 + number_of_participants - position

    if tid != None: # If the tid is None, they have not raced

        score_workbook[klass]["A" + str(position+2)] = position + 1
        score_workbook[klass]["B" + str(position+2)] = namn # Name

        if klubb != None: # Dont write "None" as the club
            score_workbook[klass]["C" + str(position+2)] = klubb # Klubb


        score_workbook[klass]["D" + str(position+2)] = tid  # Tid
        score_workbook[klass]["E" + str(position+2)] = "{:.1f}".format(19.5/(int(datetime.datetime.strptime(tid, "%M:%S").strftime("%M"))/60 + int(datetime.datetime.strptime(tid, "%M:%S").strftime("%M"))/3600))


        if klubb == "Väsby SS Triathlon": # Only Väsby Triathlon members gets a score
            score_workbook[klass]["F" + str(position+2)] = points      # points

        return position + 1
    else: # If tid was none, dont increase the position
        return position

def fill_final_results():

    # Load the final score workbook
    final_workbook = load_workbook(filename="Syratomten Total Poängställning.xlsx")

    #final_final_result_list = []
    """final_result_dict = {
        "Herr" : {},
        "Dam" : {},
        "Herr U23" : {},
        "Dam U23" : {}

    }"""

    final_result_dict = {}

    print(final_result_dict)

    for races in workbooks_created:

        # Load the individual race score workbook
        race_workbook = load_workbook(filename=races)
        print("INFO: Opened workbook " + races)

        """
        if races == "Syratomten Deltävling 1.xlsx":
            race_column = "C"
        elif races == "Syratomten Deltävling 2.xlsx":
            race_column = "D"
        elif races == "Syratomten Deltävling 3.xlsx":
            race_column = "E"
        elif races == "Syratomten Deltävling 4.xlsx":
            race_column = "F"
        elif races == "Syratomten Deltävling 5.xlsx":
            race_column = "G"
        elif races == "Syratomten Deltävling 6.xlsx":
            race_column = "H"
        elif races == "Syratomten Deltävling 7.xlsx":
            race_column = "I"
        elif races == "Syratomten Deltävling 8.xlsx":
            race_column = "J"
        elif races == "Syratomten Deltävling 9.xlsx":
            race_column = "K"
        elif races == "Syratomten Final.xlsx":
            race_column = "M"
        """



        # For each class in the race
        for race_class in workbook_sheets:
            #print("RACE_CLASS: " + race_class)

            # Reset the final_result_list befor each loop
            #final_result_list = []

            # Append the workbook sheet values to a list because it is easier to work with
            for values in race_workbook[race_class].iter_rows(min_row=2, values_only=True):
                #final_result_list.append(values)

                #print(values)
                #print(races)
                print(values[1], values[5])
                """
                final_result_dict[race_class] = {
                    values[1] : {
                        "dt1" : values[5]
                    }
                }
                """

                if race_class not in final_result_dict:
                    final_result_dict[race_class] = {}
                #final_result_dict[race_class][values[1]] = {}

                if values[1] not in final_result_dict[race_class]:
                    final_result_dict[race_class][values[1]] = {races: values[5]}
                else:
                    final_result_dict[race_class][values[1]][races] = values[5]
                #final_result_dict[race_class][values[1]][races] = values[5]

                #final_result_dict[race_class][values[1]]["dt1"] = values[5]


                #print(final_result_dict)
                #final_result_dict["Herr"]["Henrik"]["dt1"] = values[5]


        #print(final_result_dict)

        print(json.dumps(final_result_dict, indent=4, sort_keys=True))

            #print(final_result_list)
            #final_position = 2 # Start at row 2

            # for each participant in the class
        """
            for participant in final_result_list:

                # Only add a new name if it doesn't exists in the resulsts lists already
                if participant[5] != None: # If the score is not None
                    #print(participant[0], participant[5])
                    temp_final_final_result_list = []
                    temp_final_final_result_list.append(participant[1])
                    temp_final_final_result_list.append(participant[5])
                    final_final_result_list.append(temp_final_final_result_list)
                    """
        #print (final_final_result_list[0][0])
        #for stuff in final_final_result_list:
        #    print(stuff)

        #print(final_final_result_list)

        #for


        """if race_workbook[race_class]["F" + str(final_position)].value != None: # Don't include participant that dont have a score (not Väsby SS Triathlon members)

                    print(final_position-1, race_class, race_workbook[race_class]["B" + str(final_position)].value, race_workbook[race_class]["F" + str(final_position)].value)

                    final_workbook[race_class]["B" + str(final_position)] = race_workbook[race_class]["B" + str(final_position)].value # Name
                    final_workbook[race_class][race_column + str(final_position)] = race_workbook[race_class]["F" + str(final_position)].value # Score

                    final_position = final_position + 1"""

    # Save the final_workbook after all the results are transfered
    final_workbook.save(filename="Syratomten Total Poängställning.xlsx")
    print("INFO: The workbook Syratomten Total Poängställning.xlsx was saved.")





def sort_fuction(elem):
    if elem[dt]: # If the values is not None, return that value
        return elem[dt]
    else: # If the value is None return "00:00" instead. Otherwise the sort() function will try to sort None, which doesn't work
        return "00:00"

workbooks_created = []

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

            #You can't allign the whole column, only cell by cell
            #workbook[sheet]["A2"].alignment = Alignment(horizontal='left')

            print("INFO: Sheet " + sheet + " created in workbook " + str(workbook_name))

    # Remove the sheet named "Sheet", which is created by default.
    if "Sheet" in score_workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    return workbook

def create_final_results_workbook():

    # Create the workbook
    workbook = Workbook()
    final_results_workbook_name = "Syratomten Total Poängställning.xlsx"
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

            print("Sheet " + sheet + " created.")

    # Remove the sheet named "Sheet", which is created by default.
    if "Sheet" in score_workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    workbook.save(filename=final_results_workbook_name)
    print ("INFO: Workbook " + final_results_workbook_name + " was saved")


if __name__ == "__main__":

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

                # Create a new workbook for this race
                score_workbook = create_race_workbook(workbooks)


                for stuff in result_list:

                    #if stuff[KLASS] == "Herr" and number_of_participants_herr != None:
                    # Herr U23 is also counted in the Herr class
                    if stuff[KLASS] == "Herr" or stuff[KLASS] == "Herr U23":
                        position_herr = scoreboard(stuff[NAME], "Herr", stuff[KLUBB], stuff[dt], number_of_participants_herr, position_herr)

                    #elif stuff[KLASS] == "Dam" and number_of_participants_dam != None:
                    # Dam U23 is also counted in the Dam class
                    elif stuff[KLASS] == "Dam" or stuff[KLASS] == "Dam U23":
                        position_dam = scoreboard(stuff[NAME], "Dam", stuff[KLUBB], stuff[dt], number_of_participants_dam, position_dam)

                    #elif stuff[KLASS] == "Herr U23"and number_of_participants_herru23 != None:
                    if stuff[KLASS] == "Herr U23":
                        position_herru23 = scoreboard(stuff[NAME], stuff[KLASS], stuff[KLUBB], stuff[dt], number_of_participants_herru23, position_herru23)

                    #elif stuff[KLASS] == "Dam U23" and number_of_participants_damu23 != None:
                    elif stuff[KLASS] == "Dam U23":
                        position_damu23 = scoreboard(stuff[NAME], stuff[KLASS], stuff[KLUBB], stuff[dt], number_of_participants_damu23, position_damu23)

                # Increase the deltävling by one each loop
                dt = dt + 1

                # Save the content in the score workbook
                score_workbook.save(filename="Syratomten " + workbooks + ".xlsx")
                print ("INFO: Workbook Syratomten " + workbooks + ".xlsx was saved")

    # Create the final results workbook
    create_final_results_workbook()


    """
    Save the results in the final result workbook as well
    The result is only saved if it is a member of Väsby SS Triathlon
    """
    #temp_final_positition = fill_final_results(temp_final_positition, namn, klass, points)
    fill_final_results()

    #final_workbook.save(filename="Syratomten Total Poängställning.xlsx")
    #print("The workbook Syratomten Total Poängställning.xlsx was saved.")
