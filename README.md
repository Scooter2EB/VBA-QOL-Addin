The purpose of this file is to be an easy to open/use addin for Excel that simplifies some task that people _say_ are easy but are very tedious to do. So rather than having to click buttons and scroll through menus, this addin will automate certain tasks with the push of a button. _Currently it can_:

1. Save a Backup Copy of your current workbook (creates folder in file source folder call 'Backups' and saves a copy with the date and number ver. no. that day)
2. Save and close all other excel workbooks except the active one.

The rest of these task require the use of Classic Outlook, not the new Outlook for Windows (so pre-2024 outlook app). Each option will open a new email in Outlook and do what is described

3. Write the current (cell) selection to an email
4. Write the current worksheet’s print area to an email
5. Attach the current worksheet’s print area as a pdf to an email
6. Save and attach the current selection as a .xlsx to an email
7. Save and attach the current worksheet’s print area as a .xlsx to an email

More features are planned, and I am open to suggestions!

_To Use_: Download and open either before you open excel or at any point after you have opened excel. You should see a new option in the Ribbon (very top of excel) labelled QOL features. Click it to see your new buttons!

Works on Windows. I have tried to put in measure to ensure it works on Apple/Linux but I don’t know so if it does please let me know!

Note: If it asks to enable Macros, you have to allow them as this is a .xlam file. If instead of asking you to enable macros, windows informs you that macros are blocked for this file, it is due to a recent-ish security feature Microsoft places on all files downloaded from the internet and email. Simply right click on the file, select ‘Properties’, and in the bottom right hand corner select Unblock before clicking ‘Apply’.  I know it seems like a lot of hurdles, but macro files have been used to send malicious code in the past, so Microsoft is just trying to protect you from unknown files.

Code is open and unlocked in the file and contains some documentation within the code (quality depends on how close to Friday it was when I wrote it…….)
