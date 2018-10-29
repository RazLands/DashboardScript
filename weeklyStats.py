# $language = "python"
# $interface = "1.0"
import csv

def cli2file(tempfile_path, waitStrs):
    # stats list variable:
    # M-Wireless RX                 M-Wireless TX
    # [FL08,IL156,IL01,ZPL13,ZMY33],[FL08,IL156,IL01,ZPL13,ZMY33]
    # M-Guest RX                    M-Guest TX
    # [FL08,IL156,IL01,ZPL13,ZMY33],[FL08,IL156,IL01,ZPL13,ZMY33],
    stats = [
        [[0, 0, 0, 0, 0],[0, 0, 0, 0, 0]],
        [[0, 0, 0, 0, 0],[0, 0, 0, 0, 0]]
    ]

    tempfile = open(tempfile_path, 'wb')
    worksheet = csv.writer(tempfile)
    i = 0
    row = 1
    while True:
        result = crt.Screen.WaitForStrings(waitStrs)  # Wait for the linefeed at the end of each line
        if result == 3:  # We saw the prompt, we're done.
            break

        # Cut the output data from current row on screen, from char 1 to 40
        screenrow = crt.Screen.CurrentRow - 1
        readline = crt.Screen.Get(screenrow, 1, screenrow, 300)

        # Split the line ( " " delimited) and put some fields into Excel
        items = readline.strip("*").strip("-").strip(" ")#.split(" ")

        # Check if M-Guest or M-Wireless
        if str(items)[0:8] == "M-Guest " or str(items)[0:17] == "M-Guest-ZMY33-QoS":
            # Check if M-Guest TX's size is GB, then convert to TB (divide by 1000)
            if str(items)[29:31] == "GB":
                stats[1][1][i - 1] = float(str(items)[21:28]) / 1000 # TX
                # Check if M-Guest RX is GB or TB, divide by 1000 if GB
                if str(items)[42:44] == "GB":
                    stats[1][0][i - 1] = float(str(items)[34:41]) / 1000 # RX
                elif str(items)[42:44] == "TB":
                    stats[1][0][i - 1] = float(str(items)[34:41]) # RX
            # Check if M-Guest TX's size is TB
            elif str(items)[29:31] == "TB":
                stats[1][1][i - 1] = float(str(items)[21:28]) # TX
                # Check if M-Guest RX is GB or TB, divide by 1000 if GB
                if str(items)[42:44] == "GB":
                    stats[1][0][i - 1] = float(str(items)[34:41]) / 1000 # RX
                elif str(items)[42:44] == "TB":
                    stats[1][0][i - 1] = float(str(items)[34:41])  # RX
        # Check if M-Guest or M-Wireless
        if str(items)[0:11] == "M-Wireless ":
            # Check if M-Wireless TX's size is GB, then convert to TB (divide by 1000)
            if str(items)[29:31] == "GB":
                stats[0][1][i - 1] = float(str(items)[21:28]) / 1000 # TX
                # Check if M-Wireless RX is GB or TB, divide by 1000 if GB
                if str(items)[42:44] == "GB":
                    stats[0][0][i - 1] = float(str(items)[34:41]) / 1000 # RX
                elif str(items)[42:44] == "TB":
                    stats[0][0][i - 1] = float(str(items)[34:41]) # RX
            # Check if M-Wireless TX's size is TB
            elif str(items)[29:31] == "TB":
                stats[0][1][i - 1] = float(str(items)[21:28]) # TX
                # Check if M-Wireless RX is GB or TB, divide by 1000 if GB
                if str(items)[42:44] == "GB":
                    stats[0][0][i - 1] = float(str(items)[34:41]) / 1000 # RX
                elif str(items)[42:44] == "TB":
                    stats[0][0][i - 1] = float(str(items)[34:41])  # RX

        if "statistics on" in str(items): i += 1
        row = row + 1

    worksheet.writerow(stats)
    tempfile.close()
    return stats

def read_temp_file(tempfile_path):
    tempfile = open(tempfile_path, 'r')
    reader = list(csv.reader(tempfile))

    return reader

def clear_statistics():
    # crt.Screen.Send("service clear wireless wlan statistics on AMERFL08" + chr(13))
    # crt.Screen.Send("service clear wireless wlan statistics on AMERIL156" + chr(13))
    # crt.Screen.Send("service clear wireless wlan statistics on AMERIL01" + chr(13))
    # crt.Screen.Send("service clear wireless wlan statistics on EMEAZPL13" + chr(13))
    # crt.Screen.Send("service clear wireless wlan statistics on ASIAZMY33" + chr(13))
    crt.Screen.Send("sh wire ap on AMERFL08" + chr(13))

def main():
    crt.Screen.Synchronous = True  # Start CRT session
    tempfile_path = "C:\DashboardScript\\temp.csv"
    waitStrs = ["\n", "60#", "Done"]

    crt.Screen.Send("en" + chr(13))
    crt.Screen.Send("sh wire wlan statistics on AMERFL08" + chr(13))
    crt.Screen.Send("sh wire wlan statistics on AMERIL156" + chr(13))
    crt.Screen.Send("sh wire wlan statistics on AMERIL01" + chr(13))
    crt.Screen.Send("sh wire wlan statistics on EMEAZPL13" + chr(13))
    crt.Screen.Send("sh wire wlan statistics on ASIAZMY33" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("sh ver" + chr(13))
    crt.Screen.Send("Done" + chr(13))

    stats = cli2file(tempfile_path, waitStrs)

    crt.Screen.Synchronous = False
    crt.Quit()

    return stats
main()

