# HorizonLogbook - an Excel workbook + horizonlogbook.exe to synchronise the survey logbook to horizon via MQTT
### Created: April 2025
### Copyright: @ ReachSubsea


# Design Approach
1. the user runs excel to create logbook entries as he does today.  no real change there.
2. excel will auto save at a user defined interval. the horizonlogbook.exe will monitor the excel workbook on disc for changes.  
3. when a change is detected, horizonlogbook will analyse what has changed and sent the DELTA'S to the horizon postgres server via MQTT
3. there is a receiving application on the reach portal which merges the changes into postgres.
5. the horizon portal updates with the new changes to postgres and the user can see the new logbook entries
4. in certain circumstances, eg the user inserts/deletes new columns or new rows into the logbook, the system will need to send the entire worksheet.  this should not happen too often but is accommodated.

# Installation
* run the setup.exe file to install.  strongly recommend you install to the default folder
* the install also installs a template excel sheet of the reach online logbook.  you can move this to any folder you prefer, but please do not move the executables
 
# How to use the online logbook
* install using the setup.exe
* you can move the excel sheet to a different folder if it suits
* you need to know the IP address of the MQTT broker we use to publish the logbook to the horizon portal.  you can ask the ICT team for that IP
* open the excel sheet and open the SETTINGS worksheet.
* CONFIGURE the PROJECT CODE, VESSEL, ROV NAMEs.  they are marked as essential because they are required to send the correct logbook to the correct logbook in the portal.
* CONFIGURE the BROKER IP address to point to the broker on your vessel.  
* CONFIGURE the path to the horizonlogbook.exe
* RESTART the horizonlogbook sync process with the button in the settings sheet.
* whenever you open the workbook, the sync tool will auto start
* whenever you open the workbook, the worksheet will automatically set itself to auto-save at the user defined interval in the settings sheet.
* as you make changes, they will be published to the horizon portal so clients and guests can see whats happening.
* Excel has a status bar.  you can see a message written there each time we publish along with the number of cells changed.

# 2DO
* Complete the horizonlogbook.exe
* Ensure sync to postgres is working correctly
* Integrate postgres into horizon portal

# Done
* Add “DailyLog-Horizonlogbook.xlm” to installer
* Make an installable for horizonlogbook
* Add FLA to installer
* Put sources into gitlab
* Write a readme how to use horizonlogbook.
* Provide installer for testing alongside navipac
* Removed all the sqlite
* Removed the headers from the logbook
* Moved the headers to a settings sheet
* Cleaned up and removed the hidden sheet
* Removed all unrequired VBA
* Removed all VBA forms
* Tested it still works with FLA
* Tested against notepad.exe to ensure it starts when excel opens (a test/demo only)
* Tested against notepad to ensure it closes when excel closes (a test/dome only)
* Added broker IP to the settings sheet
* Added project name to the settings sheet
* Added vessel to settings sheet
* Save broker IP address to an ini file (not sure we need this)
* Added comments on how to use to the settings sheet
* Added a restart button to the sync process (more of a test feature)
* Added option to browse to the horizonlogbook.exe which syncs
* Use named cells rather than cell addresses in the settings sheet
* Simplified the Logbook to remove merged cells, redundant headers.  The logbook is now a simple of rows and columns which means we can sync with confidence and reliability.
* The VBA code is massively reduced to just some startup and to ensure the python sync tool is running.
* Save the settings to an ini file so horizonlogbook.exe knows which excel sheet we are trying to sync, the update rate etc. an ini file is easy to create, read and work with compared to permissions of an xl sheet.

![Logbook Sheet example](images/Logbook1.png)

![Settings Sheet example](images/settings1.png)

![Resource Sheet example](images/resources1.png)

![Logbook and FLS example](images/FLA1.png)
