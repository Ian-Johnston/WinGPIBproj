WinGPIB
App rebranded 2nd Jan. 2022 - UK/USA Trademark.
wingpib.co.uk & wingpib.com

This app provides reliable Windows GPIB functionality to many brands and types of test instruments, aka HPIB, GPIB, or IEEE-488 Instrument Control.

WinGPIB is for non-commercial use only, contact me if you wish to use it commercially.

Written in Visual Studio 2019 VB 2022 VB and using Pawel Wzietek's awesome GPIB library code (.DLL Multi-threaded GPIB communication), I wrote a custom front end to what you see below.
The app can interface using VISA, GBIP:ADlink, GPIB-488 and COM port......and can run two devices at the same time.
Pawel's original project: https://www.codeproject.com/Articles/1166996/Multithreaded-communication-for-GPIB-Visa-Serial-i

I/O DEVICE COMPATIBILITY:
99% of instruments with a GPIB inteface will work with this app because it is extremely flexible.
- VISA
- GBIP:ADlink
- GPIB-488
- COM port (serial)
- Prologix (Ethernet TCP & USB Serial)

The app primarily works with Keysight IO Libraries and using the Agilent E5810A or the Agilent 82357B etc. Not to say it won't work with other interfaces, your mileage may vary.

TEMPERATURE/HUMIDITY DEVICE COMPATIBILITY:
- USB-TnH SHT10 V2.00 & USB-TnH SHT30 (Usb serial comms).
- User based protocol (partial).
- Adafruit MCCP2221A/SHT40 as follows:
  Adafruit MCP2221A Breakout - General Purpose USB to GPIO ADC I2C - Stemma QT / Qwiic - http://adafru.it/4471
  Adafruit Sensirion SHT40 Temperature & Humidity Sensor - STEMMA QT / Qwiic - http://adafru.it/4885
  STEMMA QT / Qwiic JST SH 4-pin Cable - 100mm Long - http://adafru.it/4210

OS TESTED:
Win10
Win11 - reports in so far that it is ok with Win11.

INSTALLATION INSTRUCTIONS:
Fresh Installation of WinGPIB:
First, make sure you have the Keysight IO Libraries loaded, running and configured. Then,
1. Unzip and run the .msi installer - WinGPIB_V#####.zip
2. The program is installed at: C:\Program Files (x86)\IanJ\WinGPIB\WinGPIB.exe
3. All user data resides in C:\Users\[username]\Documents\WinGPIBdata
4. A desktop shortcut is created to run the app.

Updating WinGPIB app:
1. Install the same .msi installer as above. It will not overwrite the files in your data folder, C:\Users\[username]\Documents\WinGPIBdata

Moving your WinGPIB profiles to a new PC:
V3.178 onwards:
1. Use the EXPORT PROFILES button to save all your profiles/settings data to \WinGPIBdataProfilesData.dat.
2. Copy the .dat file to your new PC \WinGPIBdata folder.
3. Once you have installed WinGPIB, use the IMPORT PROFILES button to import all your profiles/settings data.
4. Restart WinGPIB.
Prior to V3.178:
1. On your old PC (I'm using Win10) go to: C:\Users\YOURUSERNAME\AppData\Local\Microsoft\WinGPIB.exe_Url_somegreatlonglistoflettersnumbers\1.0.0.0\
and in there is a file called user.config
2. Copy that file out and put it in the similarly named folder in your new PC. The large folder name will be different.

Installation Notes:
1. If you accidentally delete your GPIBchannels.txt or base Log.csv file there are original copies in the program folder you can copy over to C:\Users\[username]\Documents\WinGPIBdata.
2. When upgrading you may lose the contents of your profiles so Export them first.
3. WinGPIB does not use OneDrive for it's data storage.
4. If you get an error on running the program that a Log.csv file or its path doesn't exist, then manually create the exact path and drop a copy of Log.csv in that folder. Restart WinGPIB.
5. Version 3.143 and over requires MCP2221DLL-M-dotNet4.zip available in the downloads. Note that depending on your Windows install you may need to install Microsoft Visual C++ 2010 Service Pack 1 Redistributable Package in order for the MCP2221DLL to work properly. Available here:-
https://www.microsoft.com/en-gb/download/details.aspx?id=26999 or https://www.techpowerup.com/download/visual-c-redistributable-runtime-package-all-in-one/

Possible VISA32.DLL Issues:
If you have Keysight IO Libraries loaded and you can communicate with your device using it's GPIB address via the Connection Expert but get an error (Visa error BFFF009E) using WinGPIB and you've tried everything, then:
Go to C:\Windows\SysWOW64
Rename Visa32.dll to Visa32_old.dll
Rename visa32.Agilent Technologies - Keysight Technologies.dll to Visa32.dll
Reason: Windows updates and updates to other programs on your PC can often replace this .DLL with their own and breaks WinGPIB Visa functionality!

INSTRUCTIONS / NOTES FOR USE (may not be complete):
Assumes that you have loaded the current version of Keysight IO Libraries Suite and have it configured.

DEVICES:
Set up your NAME and ADDRESS of the device(s) you want to connect to, and pick the INTERFACE type (I have tested Visa only).
The defaults shown when you install are my own as an example.
Hit the CREATE DEVICE button (you choose). An pop-up box will appear that lists the connected device. This box will remain visible at all times and will let you see the commands being sent and any queued commands.

You can effectively batch commands using the PRE RUN, AT RUN and AT STOP boxes and using the RUN/STOP button at the bottom.
PRE RUN commandsare used to set up your device.
AT RUN command is the command that is repeated at the SAMPLE RATE.

Command types are as follows:

SINGLE COMMANDS:
Enter single commands in the COMMAND boxes and hit QUERY and you'll get a RESPONSE back.
Types of commands as follows:-
Query Blocking commands are immediately executed/sent. They are immediately executed on the calling thread (usually GUI thread), this method waits until it gets response from the interface.
Query Async commands are queued. Queries are queued and the queue is processed on a different thread (producer-consumer model). The call appends the query to the queue and returns immediately.
Send Async  commands queued. These types of commands do NOT expect a reply.

BATCH/LOOPING COMMANDS:
PRE RUN - Initialize your device with a list of SEND ASYNC type commands.
AT RUN - A single QUERY ASYNC type command.
AT STOP - A list of SEND ASYNC type commands you may wish to send when you hit STOP.

INTERRUPT COMMANDS:
You can specify an Async (no reply) command to be send periodically, such as ACAL DCV when you are logging etc.
Duration in secs = How long you need that command to be active for, i.e. ACAL DCV on a 3458 may take 60secs.
Period in mins = How often do you want that command sent, i.e. every 120mins.

IO DEVICES POP-UP:
When the CREATE DEVICE buttons are pressed a pop-up box will appear that lists the connected device(s).
This box will remain visible at all times and will let you see the commands being sent and any queued commands.
Note that you can close this window without any effect on the app. You can use the button "Show Device List" to show it again if needed.

COMMAND LINE INTERFACE:
After connecting to your device(s) use the comand line interface to send and receive data from your device(s).

TEMPERATURE/HUMIDITY:
There is only a couple of Temperature/Humidity sensor I have tested and that are available from DOGRATIAN seller here - https://www.dogratian.com/index.php
They are serial USB type sensor and does not require any drivers except what Win7/Win10 loads automatically as you insert it. You should see a COM port which you can pick from the list on the app.
The SHT30 gives 2dp's whereas the SHT10 1dp.

DATA LOG / CSV:
For CSV generation, per the defaults, set up your filename and path to where you want the CSV to be saved. At the moment there is not file generated if it doesn't exist so make sure you have an empty .CSV text file at the path specified.
The check box ENABLE CSV will start saving data to the .CSV file.
The EXPORT CSV file will copy the active .CSV file to the same folder and append the current date & time as the filename.
EXPLORER will launch Windows Explorer and pointing to the path folder.
You can clear the current CSV file with the RESET CSV button, there's a check box next to it which must be set before the reset will work.
ENABLE LOG check box will activate the log file display window.
E NOTATION TO DECIMAL check box will convert ########E-02 to the decimal equivalent. This is for the CSV and well as the log window.

CHART:
Enable either or both device charts via the check boxes.
With the boxes un-checked you can hit the CLEAR CHART to reset the chart windows and remove the data.
You can change the resolution of the chart with the X-AXIS SCALE POINTS entry. Once the chart display reaches the set number of data points on it the chart will start to scroll off to the left as new data appears at the right.
The range of the Y-axis is adjustable via the SCALE MIN AND SCALE MAX enries. You can manually set them or hit AUTO SCALE and the chart will look at the existing trace and from it's highest and lowest position will centre and scale the chart.
The UP and DN buttons allow the current set scale to be adjusted in 10% steps.

PLAYBACK CHART:
For offline display of the generated .CSV files. The user can select the Device to be displayed on the graph (since there are 2 devices available), and also zoom in/out and scroll back/forward through the graph.
TIP: The Playback Chart relies on the consistency of the CSV being correct, i.e. no empty lines, the first 3 lines start with // for the metadata, the first and last data timestamp is used to calculate the period of the entire data.

3458A & 3457A DMM RAM EXTRACT:
Utility to extract the battery backed up (DS1220Y IC) cal ram on the 3458A down to a .bin file.
Utility to extract the battery backed up (DS1235Y ICs) settings ram (both) on the 3458A down to a .bin file.
Utility to extract the battery backed up cal ram (2 versions) on the 3457A down to a .bin file.

3245A UNIVERSAL SOURCE CALIBRATION:
Utility to automate the full calibration (DCV & DCI) of an HP 3245A by using a 3458A.

SAVE SETTINGS:
Most of the entries are saved off so when you restart the app they will be there.

E5810A AGILENT LAN/GPIB GATEWAY CONFIG:
1. Power up E5810A and make sure it gets an IP assigned (DHCP), i.e. 192.168.1.175
2. Go to http://192.168.1.175 and "Find & Query Instruments". Make sure you can communicate with your device.
3. Go to Keysight Connection Expert and ADD Remote GPIB Interface.
    Change SICL Interface ID to the one that finds the IP of the E5810A. Hit OK.
4. Now a device will be listed on the "My Instruments" panel. The GPIB address will be displaye i.e. GPIB1::22:INSTR

TERMINATOR INFO:
CR carriage return……..moves to column zero without advancing paper
LF line feed…… advances paper without returning to column zero
CRLF……. does both the above
\r = CR = 0x0D = 13
\n = LF = 0x0A = 10 = newline

COMMON PROBLEMS:
ISSUE 1: MAV not set errors:
Example: "Poll timeout: MAV not set, status byte=4"
Check your device manual to see where the MAV is coded in the status byte and set the MAV mask accordingly.
The standard value is 16 and is set by default in WinGPIB.
If it still doesn't work the only option is to disable polling altogether although you then might experience
delays if you have more than one device connected via GPIB.
Also, if this is happening with large NPLC settings or other such time hungry commands then try increasing the Timeout (ms) to 10000 or 15000.

ISSUE 2: Can connect to my device but get syntax errors, no command is actioned:
I have found a dodgy GPIB cable/connector can cause this.

ISSUE3: The main WinGPIB.exe is sometimes flagged up & quarantined by Windows anti-virus scans. Three Four times now it's been flagged as containing a virus and three four times now I have sent it in and it's been confirmed as a false positive. Twice by BitDefender, and once twice by Microsoft Defender. If it happens to you then please let me know and I will contact the anti-virus company involved.

Have fun.

Ian.
