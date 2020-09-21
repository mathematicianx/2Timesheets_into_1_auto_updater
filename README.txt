This script was used to automate filling two different Project Timesheets (excel files). Since the information was doubled and it was necessary to update each excel daily by hand I wrote a script to automate this task.
First the script connects to dropbox api and downloads two excel files. Then the files are modified using openpyxl and updated file is uploaded to dropbox.
Each day there is also a backup copy made in case something goes wrong.
This script is used with Windows Task Manager or Cron on Linux (raspberry pi).