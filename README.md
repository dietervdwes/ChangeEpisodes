# ChangeEpisodes
Script to edit a predefined list of episodes in TrakCare laboratory information system

# What this does:
It reads a text file containing the episode numbers to be cleared: clear_list.csv;change_list.csv or rebill_list.csv (which must be in the same directory as the script.

# How to use:
Edit the file(s) called clear_list.csv;change_list.csv or rebill_list.csv to add the episodes numbers to be cleared.
Make sure the episode numbers are suffixed in the second column by the test set to be cleared (universal TrakCare code or synonym).  This step is not necessary for the Rebilling function.

Open Result Entry Window / Close an active episode's Result Entry - Single Window.

Run the Script: Change_Episodes.ahk
Hit whichever button on the script you need.
Do not do other actions on the PC while running.
It will execute for each episodes listed in the CSV line by line.

# Report bugs or ask questions
dieter.vdwesthuizen.nhls.ac.za

