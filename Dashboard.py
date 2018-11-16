'''
This script automating the weekly GWT dashboard presentation.
To be able to update the dashboard, the script updates the data stored in the google-shit file.
You need to share the spread shit with this Email: automation-dash@automation-dash.iam.gserviceaccount.com to grant access
'''

print(
    " *********************************************** \n" \
    " **** Welcome to Automatic Dashboard Script **** \n" \
    " *** This Script Was Coded By Raz Landsberger ** \n" \
    " *********************************************** \n"
)

# Import modules
import csv
import os
import gspread
import pivotCSV
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials
from decimal import Decimal
from datetime import date

# Update (write to) a specific range of cells in Google sheets by a given table of data
def update_sheet(ws, table, rangeStart, rangeEnd, firstRow):
    for index, row in enumerate(table):
        range = '{start}{i}:{end}{i}'.format(
            start=rangeStart, end=rangeEnd, i=index+firstRow
        )
        cell_list = ws.range(range)

        for i, cell in enumerate(cell_list):
            if "week" not in str(table[index][i]):
                table[index][i] = float(table[index][i])
            cell.value = row[i]

        ws.update_cells(cell_list)

# Update the "last 5 weeks" table and push the new week at the top
def week_update_func(startRange, endRange):
    current_week = int(str(file.values_get(startRange)["values"][0])[7:9]) + 1
    week_update = file.values_get(startRange+":"+endRange)
    i = 4
    while i > 0:
        week_update["values"][i] = week_update["values"][i - 1]
        i -= 1

    return [week_update["values"], current_week]

# Read results stored in temp.csv
def read_temp_file(tempfile_path):
    tempfile = open(tempfile_path, 'r')
    reader = list(csv.reader(tempfile))

    return reader

# use creds to create a client to interact with the Google Drive API
script_path = "C:\DashboardScript\\"
backup_file = open(f"{script_path}\Backups\\backup_{str(date.today())}.txt", "w")
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(f"{script_path}client_secret.json", scope)
client = gspread.authorize(creds)

# Find a specific worksheet by its URL address
file = client.open_by_url("https://docs.google.com/spreadsheets/d/1-NBSrRiLjPXLreCfbcBWJ2M7CEK1qG2BnRhNNMKkX4Q/")
tempfile_path = f"{script_path}temp.csv"
sheet = file.worksheet("Summary Graph")

'''Weekly usage per site'''
# Run SecureCRT script to check for sites' usage.
os.system(f"{script_path}weeklyStats.bat")
reader = read_temp_file(tempfile_path)
mg_usage = (reader[0][1].strip("[").strip("]")).split(", ")
mg_usage_rx = mg_usage[0:5]
mg_usage_tx = mg_usage[5:10]

mw_usage = (reader[0][0].strip("[").strip("]")).split(", ")
mw_usage_rx = mw_usage[0:5]
mw_usage_tx = mw_usage[5:10]

# Fix the table of usage per site (FL08, IL156, IL01, ZMY33, ZPL13) for MW & MG
usage_per_site = [
    [float(mw_usage_rx[0]), float(mw_usage_tx[0].strip("[")), float(mg_usage_rx[0]), float(mg_usage_tx[0].strip("["))], # FL08: [MW RX, MW TX, MG RX, MG TX]
    [float(mw_usage_rx[1]), float(mw_usage_tx[1]), float(mg_usage_rx[1]), float(mg_usage_tx[1])], # IL156: [MW RX, MW TX, MG RX, MG TX]
    [float(mw_usage_rx[2]), float(mw_usage_tx[2]), float(mg_usage_rx[2]), float(mg_usage_tx[2])], # IL01: [MW RX, MW TX, MG RX, MG TX]
    [float(mw_usage_rx[3]), float(mw_usage_tx[3]), float(mg_usage_rx[3]), float(mg_usage_tx[3])], # ZPL13: [MW RX, MW TX, MG RX, MG TX]
    [float(mw_usage_rx[4].strip("]")), float(mw_usage_tx[4]), float(mg_usage_rx[4].strip("]")), float(mg_usage_tx[4])] # ZMY33: [MW RX, MW TX, MG RX, MG TX]
]

# Call week_update_func to update the "Weekly usage per site" table with the new data
update_sheet(sheet,usage_per_site,'C','F',35)
print(
    "Weekly usage per site has updated\n"
    " MW RX ------ MW TX ------ MG RX ------ MG TX \n"
    f" {round(usage_per_site[0][0],2)} ------- {round(usage_per_site[0][1],2)} ------- {round(usage_per_site[0][2],2)} ------- {round(usage_per_site[0][3],2)}\n"
    f" {round(usage_per_site[1][0],2)} ------- {round(usage_per_site[1][1],2)} ------- {round(usage_per_site[1][2],2)} ------- {round(usage_per_site[1][3],2)}\n"
    f" {round(usage_per_site[2][0],2)} ------- {round(usage_per_site[2][1],2)} ------- {round(usage_per_site[2][2],2)} ------- {round(usage_per_site[2][3],2)}\n"
    f" {round(usage_per_site[3][0],2)} ------- {round(usage_per_site[3][1],2)} ------- {round(usage_per_site[3][2],2)} ------- {round(usage_per_site[3][3],2)}\n"
    f" {round(usage_per_site[4][0],2)} ------- {round(usage_per_site[4][1],2)} ------- {round(usage_per_site[4][2],2)} ------- {round(usage_per_site[4][3],2)}\n\n\n"
)

backup_file.write(
    "Weekly usage per site\n"
    " MW RX ------ MW TX ------ MG RX ------ MG TX \n"
    f" {round(usage_per_site[0][0],2)} ------- {round(usage_per_site[0][1],2)} ------- {round(usage_per_site[0][2],2)} ------- {round(usage_per_site[0][3],2)}\n"
    f" {round(usage_per_site[1][0],2)} ------- {round(usage_per_site[1][1],2)} ------- {round(usage_per_site[1][2],2)} ------- {round(usage_per_site[1][3],2)}\n"
    f" {round(usage_per_site[2][0],2)} ------- {round(usage_per_site[2][1],2)} ------- {round(usage_per_site[2][2],2)} ------- {round(usage_per_site[2][3],2)}\n"
    f" {round(usage_per_site[3][0],2)} ------- {round(usage_per_site[3][1],2)} ------- {round(usage_per_site[3][2],2)} ------- {round(usage_per_site[3][3],2)}\n"
    f" {round(usage_per_site[4][0],2)} ------- {round(usage_per_site[4][1],2)} ------- {round(usage_per_site[4][2],2)} ------- {round(usage_per_site[4][3],2)}\n\n\n"
)

'''Clients amount per site'''
# Receive MW & MG clients amount per site from CSV report
WLAN = pivotCSV.wlan_counter()
mguest = WLAN["mguest"] # Clients amount per site for MG
mwireless = WLAN["mwireless"] # Clients amount per site for MW

# Fix a table of the amount of clients per site (FL08, IL156, IL01, ZMY33, ZPL13) for MW & MG
clients_per_site =[[mwireless["FL08"], mguest["FL08"]],
                   [mwireless["IL156"], mguest["IL156"]],
                   [mwireless["IL01"], mguest["IL01"]],
                   [mwireless["ZPL13"], mguest["ZPL13"]],
                   [mwireless["ZMY33"], mguest["ZMY33"]]]
# Call week_update_func to update the "Clients amount per site" table with the new data
update_sheet(sheet, clients_per_site, 'C', 'D', 13)
print(
    "Clients amount per site has updated\n"
    ' M-Wireless --- M-Guest\n'
    f'  {clients_per_site[0][0]} ------- {clients_per_site[0][1]}\n'
    f'  {clients_per_site[1][0]} ------- {clients_per_site[1][1]}\n'
    f'  {clients_per_site[2][0]} ------- {clients_per_site[2][1]}\n'
    f'  {clients_per_site[3][0]} ------- {clients_per_site[3][1]}\n'
    f'  {clients_per_site[4][0]} ------- {clients_per_site[4][1]}\n\n\n'
)

backup_file.write(
    'Clients amount per site\n'
    ' M-Wireless --- M-Guest\n'
    f'  {clients_per_site[0][0]} ------- {clients_per_site[0][1]}\n'
    f'  {clients_per_site[1][0]} ------- {clients_per_site[1][1]}\n'
    f'  {clients_per_site[2][0]} ------- {clients_per_site[2][1]}\n'
    f'  {clients_per_site[3][0]} ------- {clients_per_site[3][1]}\n'
    f'  {clients_per_site[4][0]} ------- {clients_per_site[4][1]}\n\n\n'
)

'''Weekly usage per WLAN'''
# Receive MW & MG clients' usage from CSV report
clients_usage = pivotCSV.usage_counter()
mg_usage = clients_usage["mguest"]
mw_usage = clients_usage["mwireless"]

# Call week_update_func to update the "Weekly usage per WLAN" table with the new data
wu_usage = week_update_func("A60","E64")
wu_usage[0][0] = [f"week {wu_usage[1]}", mw_usage["RX"], mw_usage["TX"], mg_usage["RX"], mg_usage["TX"]]
update_sheet(sheet,wu_usage[0],'A', 'E', 60)
print(
    "Weekly usage per WLAN has updated\n"
    " Week #  ---- MW RX ----- MW TX ------ MG RX ------ MG TX \n"
    f" {wu_usage[0][0][0]} ---- {wu_usage[0][0][1]} ------ {wu_usage[0][0][2]} ------ {wu_usage[0][0][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][1][0]} ---- {wu_usage[0][1][1]} ------ {wu_usage[0][1][2]} ------ {wu_usage[0][1][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][2][0]} ---- {wu_usage[0][2][1]} ------ {wu_usage[0][2][2]} ------ {wu_usage[0][2][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][3][0]} ---- {wu_usage[0][3][1]} ------ {wu_usage[0][3][2]} ------ {wu_usage[0][3][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][4][0]} ---- {wu_usage[0][4][1]} ------ {wu_usage[0][4][2]} ------ {wu_usage[0][4][3]} ------- {wu_usage[0][0][4]}\n\n"
)

backup_file.write(
    "Weekly usage per WLAN\n"
    " Week #  ---- MW RX ----- MW TX ------ MG RX ------ MG TX \n"
    f" {wu_usage[0][0][0]} ---- {wu_usage[0][0][1]} ------ {wu_usage[0][0][2]} ------ {wu_usage[0][0][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][1][0]} ---- {wu_usage[0][1][1]} ------ {wu_usage[0][1][2]} ------ {wu_usage[0][1][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][2][0]} ---- {wu_usage[0][2][1]} ------ {wu_usage[0][2][2]} ------ {wu_usage[0][2][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][3][0]} ---- {wu_usage[0][3][1]} ------ {wu_usage[0][3][2]} ------ {wu_usage[0][3][3]} ------- {wu_usage[0][0][4]}\n"
    f" {wu_usage[0][4][0]} ---- {wu_usage[0][4][1]} ------ {wu_usage[0][4][2]} ------ {wu_usage[0][4][3]} ------- {wu_usage[0][0][4]}\n\n"
)

'''Weekly total clients amount'''
# Call week_update_func to update the "Weekly total clients amount" table with the new data
wu_clients = week_update_func("A81","C85")
wu_clients[0][0] = [f"week {wu_clients[1]}",int(mguest["Total"]),int(mwireless["Total"])]
update_sheet(sheet, wu_clients[0], 'A', 'C', 81)
print(
    "Weekly total clients amount has updated\n"
    " Week #  ------ M-Wireless ---- M-Guest\n"
    f" {wu_clients[0][0][0]} ------ {wu_clients[0][0][1]} -------- {wu_clients[0][0][2]}\n"
    f" {wu_clients[0][1][0]} ------ {wu_clients[0][1][1]} -------- {wu_clients[0][1][2]}\n"
    f" {wu_clients[0][2][0]} ------ {wu_clients[0][2][1]} -------- {wu_clients[0][2][2]}\n"
    f" {wu_clients[0][3][0]} ------ {wu_clients[0][3][1]} -------- {wu_clients[0][3][2]}\n"
    f" {wu_clients[0][4][0]} ------ {wu_clients[0][4][1]} -------- {wu_clients[0][4][2]}\n\n\n")

backup_file.write(
    "Weekly total clients amount\n"
    " Week #  ------ M-Wireless ---- M-Guest\n"
    f" {wu_clients[0][0][0]} ------ {wu_clients[0][0][1]} -------- {wu_clients[0][0][2]}\n"
    f" {wu_clients[0][1][0]} ------ {wu_clients[0][1][1]} -------- {wu_clients[0][1][2]}\n"
    f" {wu_clients[0][2][0]} ------ {wu_clients[0][2][1]} -------- {wu_clients[0][2][2]}\n"
    f" {wu_clients[0][3][0]} ------ {wu_clients[0][3][1]} -------- {wu_clients[0][3][2]}\n"
    f" {wu_clients[0][4][0]} ------ {wu_clients[0][4][1]} -------- {wu_clients[0][4][2]}\n\n\n"
    f"The Dashboard is now updated to week {wu_usage[1]}"
)

backup_file.close()
print("The Dashboard is now updated to week {}".format(wu_usage[1]))
messagebox.showinfo("GWT Dashboard Script By Raz","The Dashboard is now updated to week {}".format(wu_usage[1]))

# Run SecureCRT again
os.system('"C:\Program Files\VanDyke Software\SecureCRT\SecureCRT.exe"')