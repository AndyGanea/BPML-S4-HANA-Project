# BPML-S4-HANA-Project
 Simple Script for FUE Determination by Andrew Ganea

## Installation Instructions

1. Download the project folder from this Github link by clicking on the Code button in the top right and then Download Zip.
2. Extract the zip file with software like [7-Zip](https://www.7-zip.org/).
3. Press the Windows key and type in *Windows Powershell* and make sure to run it with **administrator access**.
4. In the window that shows up, type in `Set-ExecutionPolicy Unrestricted` and then press `A`.
5. Go back to where you downloaded the folder, and copy the full path to the folder (starting with C:\...).
6. In PowerShell, type in cd and then paste in the file path.
7. Paste in `.\Install_Python.ps1` and then press Enter, be patient as the script is now installing Python and the project dependencies.

## How to Generate a Simple Excel Report

1. Now that Python is installed, right-click BPML_Python_Script.py and click Open with -> Python
2. When prompted to enter your username, type in your email and press Enter
3. When prompted to enter your password, type in the password and Press enter
    + If you have 2FA/MFA enabled, you need to paste in an app password.
    + To generate an app password, go [here](https://account.activedirectory.windowsazure.com/Proofup.aspx), log in, click on **Add Sign-in method** and choose **App Password**.
    + Once generated, keep your app password in a secure place since it will only ever be displayed once.
4. If your details are typed in correctly, the current BPML will show up, press 1 and then press enter.
5. All of the sheets will be displayed. Type in the number of the sheet that has information on the roles that are assigned to each user in the form of a grid of X's and press enter.
6. The sheets are now displayed a second time. Now, type in the number of the sheet that shows whether a role is Advanced, Core, or Self-Service and press enter.
7. The report is now automatically generated and stored in (Your Project Folder) -> spreadsheets -> output
8. The report is timestamped with the current date and time in your's computer local time zone.
