import sys,os
import glob
import shutil
import csv
from datetime import datetime
import codecs

''' 
This python program reforms the attendee CSV downloaded from Eventbrite to create a reader-friendly statistics result CSV.

The following files shall be copied under the same directory with this program.
1. Previously generated statistics result CSV(s) (if any), named "eb_statistics_result_yyyymmdd_hhmm.csv".
2. Attendee CSV(s) downloaded from Eventbrite, named "report-yyyymmdd-Thhmm.csv"
'''

# Attributes (Modify Accordingly)
att0 = ["Order No.","Order Date and Time","Name","Email"] # "Must have" attributes in every eventbrite event.
att1 = ["Require Visa","Symposium","Symposium Dinner","Pre-Event","Tech-Tour"] # Tickets.
att2 = ["Confirmation Email Sent","Reminder1 Email Sent","Reminder2 Email Sent","Reminder3 Email Sent","Check-In","Feedback Email Sent","Reserved1","Reserved2","Reserved3"] # User define area.
att4 = ["Flag"] # System message display.

# Find CWD
cwd = os.getcwd()

# Find the previously statistics result CSVs (if any). The latest file is selected.
print("")
old_statistics_flag = 0
for file in glob.glob("eb_statistics_result_2[0-9][0-9][0-9][0-1][0-9][0-3][0-9]_[0-2][0-9][0-5][0-9].csv"):
    old_statistics_flag += 1
    if old_statistics_flag == 1:
        file_split = file.split('_')
        old_statistics_datetime = int(file_split[3]+file_split[4][0:4])
        old_statistics_name = file
    else:
        file_split = file.split('_')
        old_statistics_datetime_multi = int(file_split[3]+file_split[4][0:4])
        if old_statistics_datetime_multi > old_statistics_datetime:
            old_statistics_datetime = old_statistics_datetime_multi
            old_statistics_name = file
if old_statistics_flag == 0:
    print("No previously generated statistics recult CSV is found in CWD.")
elif old_statistics_flag == 1:
    print("One previously generated statistics result CSV is found in CWD.")
    print("The previously generated statistics result CSV file name is: ",old_statistics_name)
    print("The previously generated statistics result CSV file date and time is: ",old_statistics_datetime)
else:
    print("Multiple previously generated statistics result CSVs are found in CWD. The latest statistics result CSV is selected.")
    print("The selected statistics result CSV file name is: ",old_statistics_name)
    print("The selected statistics result CSV file date and time is: ",old_statistics_datetime)

# Note the new attendee CSV downloaded from Eventbrite
print("")
new_attendee_flag = 0
for file in glob.glob("report-2020-[0-1][0-9]-[0-3][0-9]T[0-2][0-9][0-5][0-9].csv"):
    new_attendee_flag += 1
    if new_attendee_flag == 1:
        file_split = file.split('-')
        new_attendee_datetime = int(file_split[1]+file_split[2]+file_split[3][0:2]+file_split[3][3:7])
        new_attendee_name = file
    else:
        file_split = file.split('-')
        new_attendee_datetime_multi = int(file_split[1]+file_split[2]+file_split[3][0:2]+file_split[3][3:7])
        if new_attendee_datetime_multi > new_attendee_datetime:
            new_attendee_datetime = new_attendee_datetime_multi
            new_attendee_name = file
if new_attendee_flag == 0:
    print("No attendee CSV downloaded from Eventbrite is found in CWD.")
    print("The updating process is terminated.")
    print("")
    debug = input("Press any key to exit the program.")
    exit()
elif new_attendee_flag == 1:
    print("One attendee CSV is found in CWD.")
    print("The attendee CSV file name is: ",new_attendee_name)
    print("The attendee CSV file date and time is: ",new_attendee_datetime)
else:
    print("Multiple attendee CSVs are found in CWD. The latest statistics result CSV is selected.")
    print("The attendee CSV file name is: ",new_attendee_name)
    print("The attendee CSV file date and time is: ",new_attendee_datetime)

# Compare old statistics result CSV and new attendee CSV date and time
if (old_statistics_flag > 0) and (old_statistics_datetime >= new_attendee_datetime):
    print("")
    print("The selected statistcs result CSV time stamp is more or equaly recent than that of the selected attendee CSV (indicating that the attendee CSV is not up to date).")
    print("Please download the most up-to-date CSV files from Eventbrite.")
    print("The updating process is terminated.")
    print("")
    debug = input("Press any key to exit the program.")
    exit()
##
# Create a new blank CSV
print("")
new_tempstatistics_name = "eb_statistics_result_" + str(new_attendee_datetime)[0:8] + "_" + str(new_attendee_datetime)[8:12] + "_temp.csv"
try:
    os.remove(new_tempstatistics_name)
except:
    pass
print("Creating a temporary CSV to reform data from the selected attandee CSV.")
print("The temporary CSV is named ",new_tempstatistics_name)
with open(new_tempstatistics_name,'w',newline='',encoding='utf-8') as objWriteCSV:
    csvWriteCSV = csv.writer(objWriteCSV)
    statistics_title = att0 + att1 + att2 + att3 + att4
    csvWriteCSV.writerows([statistics_title])
    sum_of_attendee = [0]*len(att1)
    
    with open(new_attendee_name,'rt',encoding='utf-8') as objReadCSVAttendee:
        # Put the entire data in list
        csvReadCSVAttendee = [row for row in csv.reader(objReadCSVAttendee,delimiter=',')]
        csvReadCSVAttendee_length = len(csvReadCSVAttendee)
        
        row_read_index = 0
        current_order_number = ""
        for row_read in csvReadCSVAttendee:
            if row_read_index == 0:
                pass
            else:
                if row_read[0] == current_order_number:
                    pass
                else:
                    # New order
                    current_order_number = row_read[0]
                    current_order_datetime = row_read[1][0:4]+row_read[1][5:7]+row_read[1][8:10]+ " " + row_read[1][11:13]+row_read[1][14:16]
                    current_order_index = 0
                    # Create a blank dictionary for this order.
                    # The dictionary is used to store the infromation of an individual guest.
                    # The key of this dictionary is the name of the guest
                    # The value of this dictionary is a list of tickets information.
                    current_order_dict = dict()
                    while csvReadCSVAttendee[row_read_index+current_order_index][0]==current_order_number:
                        # Still in the order:
                        current_order_index_attendeename = (csvReadCSVAttendee[row_read_index+current_order_index][2].strip()).title() + " " + (csvReadCSVAttendee[row_read_index+current_order_index][3].strip()).title()
                        if not(current_order_index_attendeename in current_order_dict):
                            # This row is associated with a new guest. Register his/her name and email in the dictionary.
                            current_order_dict[current_order_index_attendeename] = [csvReadCSVAttendee[row_read_index+current_order_index][4],[0]*len(att1)]
                        # Register the event of the attendee.
                        if csvReadCSVAttendee[row_read_index+current_order_index][6] == "Ticket0": # Modify to the ticket name.
                            current_order_dict[current_order_index_attendeename][1][0] = 1 # "Primary Attendee","Accompany","Require Visa", "Symposium","Symposium Dinner","Pre-Event","Tech-Tour"
                            sum_of_attendee[0] += 1
                        elif csvReadCSVAttendee[row_read_index+current_order_index][6] == "Ticket1":
                            current_order_dict[current_order_index_attendeename][1][1] = 1
                            sum_of_attendee[1] += 1
                        elif csvReadCSVAttendee[row_read_index+current_order_index][6] == "Ticket2": 
                            current_order_dict[current_order_index_attendeename][1][2] += 1
                            sum_of_attendee[2] += 1
                        elif csvReadCSVAttendee[row_read_index+current_order_index][6] == "Ticket3": 
                            current_order_dict[current_order_index_attendeename][1][3] = 1
                            sum_of_attendee[3] += 1
                        elif csvReadCSVAttendee[row_read_index+current_order_index][6] == "Ticket4": 
                            current_order_dict[current_order_index_attendeename][1][4] = 1
                            sum_of_attendee[4] += 1
                        if row_read_index+current_order_index+1 == csvReadCSVAttendee_length:
                            # The attendee CSV is comming to the end. This must be the last row.
                            break
                        else:
                            current_order_index += 1
                    # All info about this order is registered in the dictionary. Write to CSV.
                    for current_order_dict_key in current_order_dict:
                        if current_order_dict[current_order_dict_key][1][0] == 0:
                            current_order_dict[current_order_dict_key][1][1] = 1
                            sum_of_attendee[1] += 1
                        current_order_dict_key_info1 = [current_order_number,
                                                       current_order_datetime,
                                                       current_order_dict_key,
                                                       current_order_dict[current_order_dict_key][0]]
                        current_order_dict_key_info2 = current_order_dict[current_order_dict_key][1]
                        current_order_dict_key_info3 = [""]*len(att3) # att3 is left blank. These blanks can be filled manually.
                        current_order_dict_key_info4 = ["New"] # The system message in flag attribute is set to "New".
                        current_order_dict_key_info = current_order_dict_key_info1 + current_order_dict_key_info2 + current_order_dict_key_info3 + current_order_dict_key_info4
                        
                        csvWriteCSV.writerows([current_order_dict_key_info])
            row_read_index += 1
    objReadCSVAttendee.close()
    csvWriteCSV.writerows([["Sum","","","",str(sum_of_attendee[0]),str(sum_of_attendee[1]),str(sum_of_attendee[2]),str(sum_of_attendee[3]),str(sum_of_attendee[4])]]) # Modify the length accordingly.
objWriteCSV.close()
print("A new temporary CSV to reform data from the selected attandee CSV has been created.")

if old_statistics_flag > 0:
    print("")
    print("Since a statistic result CSV generated previously has meen setected, the new temporary CSV is to be merged with the previously generated statistic result CSV.")
    new_statistics_name = "eb_statistics_result_" + str(new_attendee_datetime)[0:8] + "_" + str(new_attendee_datetime)[8:12] + ".csv"
    try:
        os.remove(new_statistics_name)
    except:
        pass
    # Previous statistic result CSV: old_statistics_name
    # Newly generated temporary CSV: new_tempstatistics_name
    print("Creating new statistic result CSV " +new_statistics_name)
    with open(new_statistics_name,'w',newline='',encoding='utf-8') as objWriteCSV:
        csvWriteCSV = csv.writer(objWriteCSV)
        with open(old_statistics_name,'rt',encoding='utf-8') as objReadCSVOldStatistics:
            csvReadCSVOldStatistics = [row for row in csv.reader(objReadCSVOldStatistics,delimiter=',')]
            with open(new_tempstatistics_name,encoding='utf-8') as objReadCSVNewStatistics:
                csvReadCSVNewStatistics = [row for row in csv.reader(objReadCSVNewStatistics,delimiter=',')]
                statistics_title_revised = csvReadCSVOldStatistics[0]
                # The title is copy-and-paste from the previous statistics CSV.
                csvWriteCSV.writerows([statistics_title_revised])
                
                # A 2-layer dictionary is created to hold the information of the new temporary CSV.
                # The rows from the previous statistics CSV are "searched" inside the dictionary.
                # Order No. and Name are used as the primary key for the searching for the 2 layers of the dictionary, respectively.
                # If order is not found, it means that the order is canceled and/or the name is removed.
                # If order is found but name is not found, the person is deleted or changed name.
                # If order and name are found, check whether content is changed.
                
                # Go through csvReadCSVNewStatistics to create the dictionary.
                order_dict = dict()
                row_read_index = 0
                for row_read in csvReadCSVNewStatistics[0:-1]:
                    if row_read_index == 0:
                        row_read_index += 1
                        continue
                    if row_read[0] in order_dict:
                        # This order already exists
                        if row_read[2] in order_dict[row_read[0]]:
                            # This person already exists (which should not happen at this stage)
                            pass
                        else:
                            # Add this person's info.
                            order_dict[row_read[0]][row_read[2]] = row_read # The entire line is input.
                    else:
                        # This order is not yet in the dictionary
                        # Creat the keys in 2 layer dictionary
                        order_dict[row_read[0]]={row_read[2]:row_read}
                    row_read_index += 1
                
                # Cross Check of previous statistic result CSV using the new temporary CSV
                print("")
                row_read_index = 0
                for row_read in csvReadCSVOldStatistics[0:-1]:
                    if row_read_index == 0:
                        row_read_index += 1
                        continue
                    # Check whether the order still exists in the 1st layer.
                    if row_read[0] in order_dict:
                        # Check whether the name still exists in the 2nd layer.
                        if row_read[2] in order_dict[row_read[0]]:
                            # Check content (Only related to personal information and attending events)
                            if order_dict[row_read[0]][row_read[2]][0:11] == row_read[0:11]:
                                # Basic information unchanged
                                order_dict[row_read[0]][row_read[2]] = row_read
                                order_dict[row_read[0]][row_read[2]][-1] = "Unchanged"
                            else:
                                 # Basic information is changed.
                                 print("Basic information of Order No. "+row_read[0]+" name "+row_read[2]+ " has been changed.")
                                 order_dict[row_read[0]][row_read[2]] = row_read
                                 order_dict[row_read[0]][row_read[2]][-1] = "Changed"
                        else:
                            print("A person has been removed/changed name in Order No. "+row_read[0]+ 
                                  " There was a name "+row_read[2]+ "that existed in "+old_statistics_name+ "but is not in the latest attendee list anymore.")
                    else:
                        print("An order existed in the previous statistics result CSV has been deleted. The order number was "+row_read[0]+" in "+old_statistics_name)
                    row_read_index += 1
            
                # Write the dictionary to file. The sequence of rows shall follow the new CSV.
                row_read_index = 0
                for row_read in csvReadCSVNewStatistics[0:-1]:
                    if row_read_index == 0:
                        row_read_index += 1
                        continue
                    if order_dict[row_read[0]][row_read[2]][-1] == "0":
                        order_dict[row_read[0]][row_read[2]][-1] = "New"
                    csvWriteCSV.writerows([order_dict[row_read[0]][row_read[2]]])
                    row_read_index += 1
                
                # Add statistics.
                csvWriteCSV.writerows([["Sum","","","",str(sum_of_attendee[0]),str(sum_of_attendee[1]),str(sum_of_attendee[2]),str(sum_of_attendee[3]),str(sum_of_attendee[4])]]) # Modify accordingly.
            objReadCSVNewStatistics.close()
        objReadCSVOldStatistics.close()
    objWriteCSV.close()
    try:
        os.remove(new_tempstatistics_name)
    except:
        pass
else:
    new_statistics_name = "eb_statistics_result_" + str(new_attendee_datetime)[0:8] + "_" + str(new_attendee_datetime)[8:12] + ".csv"
    print("")
    print("Since no existing statistic result CSV has meen detected, the new temporary CSV is to be renamed to "+new_statistics_name)
    os.rename(new_tempstatistics_name,new_statistics_name)

print("")
print("New statistics result CSV "+new_statistics_name+" has been created in "+cwd)
print("")
debug = input("Press Enter to exit the program.")
