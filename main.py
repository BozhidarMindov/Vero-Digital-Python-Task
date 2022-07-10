import openpyxl
import pandas
import requests
from datetime import datetime
from openpyxl.styles import PatternFill, Font
import math

API = "https://api.baubuddy.de/dev/index.php/v1/vehicles/select/active"
COLOR_API = "https://api.baubuddy.de/dev/index.php/v1/labels/"

#making a request to the API for data
response = requests.get(API).json()

#filtering API responses, so that only the json items that have an "hu" value are processed
filtered_response = []
for item in response:
    if item["hu"]:
        filtered_response.append(item)


color_codes = []
#getting the color code for each json item
for item in filtered_response:
    if item["labelIds"]:
        color_response = requests.get(f"{COLOR_API}{item['labelIds']}").json()
        color_codes.append(color_response[0]['colorCode'].replace("#", ""))
    else:
        color_codes.append("")


#since the keys in every json item in the response are the same, we can save their names and use them later to check for valid user input
keys = [key.lower() for key in filtered_response[0]]


#asking for the names of additional columns to be printed
additional_column_names = []
while True:
    k = input("Pass the names of the additional columns you want to print (Type 'ALL' to print all columns. "
              "Type 'STOP' to end): ").lower()
    #if the user types stop the loop breaks
    if k == "stop":
        break

    if k == "all":
        additional_column_names = keys
        print("All data columns will be printed!")
        break

    #if 'k' is already in the list of columns to be displayed, the last input is skipped
    if k in additional_column_names:
        print("You already requested that. Try again!")

    # if the user types some invalid input, they will be notified and prompted to type again
    elif k not in keys:
        print("Inavid input. Try again!")

    #the value is added to the list of additional columns
    #That is when the input is valid and the value is not already in the list of additional columns
    else:
        additional_column_names.append(k)
        print("Successfully added column")

#asking if the user if the columns in the excel fil should be colored
c = bool()
color_input = input("Do you want the table to be colored. Type 'yes' or  'no'.(Value is defaulted to true): ").lower()
if color_input == "no":
    c = False
else:
    c = True


#creating a list of columns that would later be placed in the xlsx file
columns_to_print = []
if "rnr" in additional_column_names:
    pass
else:
    columns_to_print.append("rnr")

if "gruppe" in additional_column_names:
    pass
else:
    columns_to_print.append("gruppe")

for item in filtered_response[0]:
    if item.lower() in additional_column_names:
        columns_to_print.append(item)


#getting the current date and iso-formating it
today = datetime.now()
current_date_iso_formatted = today.isoformat().split("T").pop(0)


#placing the data we got from the API in an excel file, sorted by the column "gruppe"
data = pandas.DataFrame(filtered_response)
data.sort_values(by="gruppe", inplace=True)

#---------------------Formating the 'hu' date, because we will need it later to determine the color of each row-------------------------

#storing the date values into an array
date_values = []
for (columnName, columnData) in data.iteritems():
    if columnName == "hu":
        date_values.append(columnData.values)


#converting the array into a string so we can replace the unneeded values
dates_str = ' '.join([str(elem) for elem in date_values])


#converting the string back to a list, whilst repalcing the unneeded values
new_date_values = list(dates_str.replace("[", "").replace("]", "").replace("'", "").replace("\n", "").split(" "))
today_date = current_date_iso_formatted.replace("-", "")

final_date_values = []
for date in new_date_values:
    new_date_value = datetime.strptime(date, '%Y-%m-%d').isoformat().split("T").pop(0).replace("-", "")
    final_date_values.append(int(today_date) - int(new_date_value))


#function that is used to determine the color of each row in the table for teh API data
def apply_color(value) -> str:
    color_fill = str()
    if value < 300:
        color_fill = "007500"
    elif value < 10000:
        color_fill = "FFA500"
    else:
        color_fill = "b30000"

    return color_fill


# --------------------------------Reading from the csv file--------------------------
data_vehicles = pandas.read_csv('vehicles.csv', delimiter=";")
data_vehicles.sort_values(by="gruppe", inplace=True)
data_dict = data_vehicles.to_dict(orient="list")

# getting the color codes of each labelId
vehicle_color_codes = []
for item in data_dict["labelIds"]:
    if math.isnan(item):
        vehicle_color_codes.append("")
    else:
        color_response = requests.get(f"{COLOR_API}{str(item)}").json()
        vehicle_color_codes.append(color_response[0]['colorCode'].replace("#", ""))



#---------------Opening an Excel Writer to place the dataframes into an excel file------------------------------
with pandas.ExcelWriter(f"vehicles_{current_date_iso_formatted}.xlsx", engine="openpyxl") as writer:
    #dropping the data that the user didn't request
    for key in filtered_response[0]:
        if key not in columns_to_print:
            data.drop(key, axis=1, inplace=True)

    #actually saving the data to an excel file
    sheet_name_api = "API Data Sheet"
    data.to_excel(writer, index= 0, sheet_name=sheet_name_api)


    #saving the data from the csv in a separate sheet
    sheet_name_vehicles = "CSV Data Sheet"
    data_vehicles.to_excel(writer, index= 0, sheet_name=sheet_name_vehicles)

#-------------------Reading from the file and coloring the rows, if the user chose that option--------------------
if c:
    #laoding a workbook and getting the needed worksheet
    wb = openpyxl.load_workbook(f"vehicles_{current_date_iso_formatted}.xlsx")
    ws = wb[sheet_name_api]

    #coloring each row of based on the "hu" value in it
    for (row, i) in zip(ws.iter_rows(min_row=2, max_col=len(data.columns), max_row=len(data) + 1), range(len(final_date_values))):
        for cell in row:
            cell.fill = PatternFill("solid", start_color=apply_color(final_date_values[i]))


    # laoding a workbook and getting the needed worksheet
    ws2 = wb[sheet_name_vehicles]

    #coloring rows where a color code for labelIds is provided and resolved
    for (row, i) in zip(ws2.iter_rows(min_row=2, max_col=len(data_vehicles.columns), max_row=len(data_vehicles) + 1), range(len(vehicle_color_codes))):
        for cell in row:
            if vehicle_color_codes[i] == "":
                pass
            else:
                cell.font = Font(color=f"{vehicle_color_codes[i]}", italic=True)

    #finally saving the excel file
    wb.save(f"vehicles_{current_date_iso_formatted}.xlsx")



