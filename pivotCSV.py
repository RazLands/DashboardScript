import csv
import tkinter as tk
from tkinter import filedialog, messagebox

tk.Tk().withdraw()
messagebox.askokcancel("GWT Dashboard Script By Raz","Please Choose The Clients CSV File Path")
client_path = filedialog.askopenfilename()
# client_path = 'C:\\Users\ptr748\Desktop\Dashboard\Client_List_5bcc4d983e487c094e9a12d3.csv'
messagebox.askokcancel("GWT Dashboard Script By Raz","Please Choose The Usage CSV File Path")
usage_path = filedialog.askopenfilename()
# usage_path = 'C:\\Users\ptr748\Desktop\Dashboard\\Usage_5b7bc2f40267fc09544cab5a.csv'
client_file = open(client_path, "r", encoding="utf8")
usage_file = open(usage_path, "r", encoding="utf8")
client_reader = csv.DictReader(client_file)
usage_reader = csv.DictReader(usage_file)

# Count clients' BYTES usage worldwide for MW & MG
usage = {"Total": {"RX": 0, "TX": 0}, "mwireless": {"RX": 0, "TX": 0}, "mguest": {"RX": 0, "TX": 0}}
def usage_counter():
    # Count the total TX and RX BYTES (MW+MG)
    for row in usage_reader:
        usage["Total"]["TX"] += float(row["TX Bytes"])/(10**12)
        usage["Total"]["RX"] += float(row["RX Bytes"])/(10**12)

    # Calculate MWireless and MGuest TX&RX BYTES (MW: 75%, MG: 25%)
    usage["mwireless"]["TX"] = float(usage["Total"]["TX"] * 0.75)
    usage["mguest"]["TX"] = float(usage["Total"]["TX"] * 0.25)
    usage["mwireless"]["RX"] = float(usage["Total"]["RX"] * 0.75)
    usage["mguest"]["RX"] = float(usage["Total"]["RX"] * 0.25)

    return usage

# Count clients amount at IL01, IL156, FL08, ZMY33, ZPL13 for both MW & MG
WLAN = {"mguest": {"Total": 0, "FL08": 0, "IL156": 0, "IL01": 0, "ZPL13": 0, "ZMY33": 0},
        "mwireless": {"Total": 0, "FL08": 0, "IL156": 0, "IL01": 0, "ZPL13": 0, "ZMY33": 0}}
mguest = ["M-Guest", "M-Guest-ZMY33-QoS", "M-Guest-QoS"]
mwireless = ["LAB-Wireless", "M-Wireless"]
def wlan_counter():
    for row in client_reader:
        if row["Wlan"] in mguest and "0.0.0" not in row["IPv4 Address"]:
            # Total MGuest clients
            WLAN["mguest"]["Total"] += 1

            # FL08 MGuest clients
            if row["RF-Domain"] == "AMERFL08":
                WLAN["mguest"]["FL08"] += 1

            # IL156 MGuest clients
            if row["RF-Domain"] == "AMERIL156":
                WLAN["mguest"]["IL156"] += 1

            # IL01 MGuest clients
            if row["RF-Domain"] == "AMERIL01":
                WLAN["mguest"]["IL01"] += 1

            # ZMY33 MGuest clients
            if row["RF-Domain"] == "ASIAZMY33":
                WLAN["mguest"]["ZMY33"] += 1

            # ZPL33 MGuest clients
            if row["RF-Domain"] == "EMEAZPL13":
                WLAN["mguest"]["ZPL13"] += 1


        elif row["Wlan"] in mwireless and "0.0.0" not in row["IPv4 Address"]:
            # Total MWireless clients
            WLAN["mwireless"]["Total"] += 1

            # FL08 MWireless clients
            if row["RF-Domain"] == "AMERFL08":
                WLAN["mwireless"]["FL08"] += 1

            # IL156 MWireless clients
            if row["RF-Domain"] == "AMERIL156":
                WLAN["mwireless"]["IL156"] += 1

            # IL01 MWireless clients
            if row["RF-Domain"] == "AMERIL01":
                WLAN["mwireless"]["IL01"] += 1

            # ZMY33 MWireless clients
            if row["RF-Domain"] == "ASIAZMY33":
                WLAN["mwireless"]["ZMY33"] += 1

            # ZPL33 MWireless clients
            if row["RF-Domain"] == "EMEAZPL13":
                WLAN["mwireless"]["ZPL13"] += 1

    return WLAN

# print(wlan_counter())