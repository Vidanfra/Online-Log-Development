| Date       | Feature                       | Description                                    | Status      |
| :--------- | :---------------------------- | :--------------------------------------------- | :---------- |
| 16/06/2025 | Remove dependency on GUID code | In order to get an smoother integration with the Horizon Log, we should remove all our Excel-SQLite synchronization logic based on the GUID. The new sync logic must be based on the Excel row position | **To Do** |
| 15/06/2025 | Fix bug when there is not TXT file for the button | When there is not a valid TXT file assigned to a button, when you try to log a new line, it doesn't write any time or date. In addition, when you press a button to log a new line, it overwrites the non TXT assigned line | **To Do** |
| 15/06/2025 | Fix bug static data reading blocking | Sometimes when you press a button to log a new line, to script freezes trying to read the static data in the Excel workbook. I think this bug normally happens when you have a Excel dialog opened or similar irregular situations. If you close and open the workbook and restart the script the bug disappears| **To Do** |
| 05/06/2025 | Hotkeys                       | Keyboard shortcuts for various functions       | **To Do**       |
| 12/06/2025 | Add Up and Down buttons in the Custom Buttons list | This would make easier for the user to reorganize the position of the buttons in the layout | **To Do** |
| 12/06/2025 | Fix bug with selected buttons Data Columns | When a row is added. moved up or down or modified. the REMOVE button makes bigger to indicate it has been selected. The problem is that the button keeps this size even when it is no longer selected. This makes the layout ugly and confusing  | **To Do** |
| 09/06/2025 | Automatic updates defined by user based on value in text file | script should be able to look at the text file and create an event when a value is reached, for example ROV depth at 15m | **Requeriment needs to be clarified** |
| 12/06/2025 | Excel Log Data | The user should define in the Data Columns menu a command to look for a certain field in the Excel Log wb. For example: Project Manager -> =SHEET2,F4 (Sheet 2, Cell F4)  | DONE |
| 12/06/2025 | Fix the bug with DB Sync | Something was broken in the Excel Log - SQL DB review process. Needs to be repaired  | DONE |
| 12/06/2025 | Reorganize button layout | Now the buttons distribuite in columns in the GUI according to the order in the Settings list. They should distribuite in rows | DONE |
| 12/06/2025 | Vehicle input field in File path menu | This will assign a value (Vessel, ROV, etc) for the KP Ref column  | DONE |
| 12/06/2025 | Adapt settings for Igor Log Excel Template| Map the new headers in the Data Columns Menu  | DONE |
| 12/06/2025 | Sequential number in Data Column tab | Display in the left a fix column of numbers to indicate the positions of the columns in the TXT file  | DONE |
| 05/06/2025 | Hourly KP in Excel            | Checkbox and automatic thread (also for new day)                                   | DONE      |
| 05/06/2025 | Save Button Sets              | Load configuration                             | DONE      |
| 05/06/2025 | TXT Files in Tabs             | One tab for each imported TXT file             | DONE      |
| 05/06/2025 | Custom Vehicle Buttons        | Buttons organized in one tab for each vehicle  | DONE       |
| 05/06/2025 | Data Headers                  | Show data from the loaded text file in headers | DONE       |
| 05/06/2025 | Auto-Open Settings            | Settings window opens automatically on startup | DONE       |
| 05/06/2025 | Layout Shape                  | Long and thin to fit in the Excel top pane     | DONE      |
| 06/06/2025 | Logger always over Excel      | Keep the logger menu always in th front        | DONE        |
| 07/06/2025 | Solve time mismatch      | Use always PC time        | DONE       |
|09/06/2025 | Colours of buttons | Should match colour set for excel row highlight | DONE |
| 09/06/2025 | Button Text colour | ability to change text colour for row in excel and in the button layout| DONE |
| 09/06/2025 | Right Click main buttons | Log on, Log off etc should be managed by right clicking rather than in settings window | DONE |
| 09/06/2025 | additional attribute added for each button for time coding | used for future DPR integration perhaps | DONE |
| 09/06/2025 | error checking if no text file associated with button | Hint to user where the error is | DONE |
| 10/06/2025 | Field Log Viewer   |   Ability to sync SQl DB to Field log viewer for offline | DONE |
| 10/06/2025 | TXT, Excel Log, SQLite Headers mapping | User defined matching the columns from the columns in every file from the Settings menu| DONE |
| 13/06/2025    |   GUID Duplication    |   If dupilicate or missing guid is found then prompt to generate new guid for conflicting cells | DONE   |