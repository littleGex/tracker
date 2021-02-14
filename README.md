# Tracker
Day time tracker

## Programming language: 
Python 3.7 and pyqt5

Set environment: XYZ

##Start GUI: 
python tracking.py or using spyder

## Document location: 
The program creates a 'Tracker' folder in C:/Users//Documents/ folder. Saving and restoring sessions, saving csv and pdf all direct to this folder as default.

When starting the GUI the user will be asked if they wish to restore a previous session to continue tracking. If no is selected a blank session is created.

## Using the timer: 
Type project (current topic) in the text field and through 'START' and 'STOP' the user operates the time tracking feature.  The 'START' function also starts when pressing ENTER whilst in the text field.

One can remove rows using the 'REMOVE ROW' button or export a csv of the table view and an excel file covering tableview, pie chart and filtered bar chart dataframes using 'Ctrl+C', using File menu or using 'EXPORT' button.

### Presenting results:
When pressing 'STOP' the date, project, start, end and time on project are added to the table view.  Automatically, but only for the current day, a pie chart appears at the bottom of the 'Tracker' tab to see what the current day time split was.  Additionally on the 'Bar chart' tab the complete time period in the table view is shown.

One can manually update the table view, whilst this will lead to update of bar chart, the 'Time on project' will not (currently) be updated - Improvement pending.

The drop-down boxes on the 'Bar chart' tab allow for filtering of dates as required (check to be added to ensure 'To' date does not preceed the 'From' date - TBD), the filter is made current by pressing the 'FILTER' button.

Additionally the bar chart can be  can be exported using the keyboard shortcut using "ctrl+P" as PDF or using the file menu options.

### Sessions: The default session shows a blank table. 
A new session can be started using "ctrl+N" or using the file menu. The session can be saved using "ctrl+S" and restored using "ctrl+R" or using Session file menu. Upon quiting ("ctrl+Q"), the user is asked whether they wish to save the session - it is not automatically saved, but this could be a default solution in a newer version.
