# ColorGoogleCalendarEvents
Spreadsheet with Google Script App that enables to automatically color events
This enables user to automtically color events on google calendar that meet certain criter (title include certain word)
The Spreadsheet itself includes an explanation, the settings the user need to configure and also menu options to initiailize and run the script
To get this solution on your account you can do two things:
1. Use this url [https://docs.google.com/spreadsheets/d/1308HNWdLj8Xdm73l2e6ytR7T_itsUrlTuVOKHRbyJ9I/copy](https://docs.google.com/spreadsheets/d/18CHrFnT3KsZ83dY6ui5FYkTvv4SC7F4XsUNKG9ju3oo/copy) and make a copy for you
![image](https://user-images.githubusercontent.com/24392647/202052041-dfe13bde-0758-4f9c-bff6-98646dd098b0.png)
2. Create a new google sheet, and enter the code under Extensions > Apps Script (You need to be familiar with App Script)

To use the solution once you did one of above
1. The sheet should open automtically after you made the copy. Wait around 15 seconds for the menu to populate
2. (First Time) Select Menu SCRIPT_STARTER > Click to Authorize. (You should only need to do steps 2-6 once ever)
![image](https://user-images.githubusercontent.com/24392647/202052294-bc4420c4-5312-49c4-b565-c489b7397a36.png)
3. Click continue to authorize
![image](https://user-images.githubusercontent.com/24392647/202052343-42f064e4-c5df-4581-9957-586648863392.png)
4. Select your google account
5. You will get a prompt that the app has not been verified. Sorry don't have plans to "verify it". The script is not malicious and code is posted here. Select Advanced, then "Go to ColorToCalendarScript"
![image](https://user-images.githubusercontent.com/24392647/202052409-4680832a-ceb7-4961-9dc1-51e4f727aed8.png)
6. You will get prompted to authorize. Do so
![image](https://user-images.githubusercontent.com/24392647/202052868-0278292c-34c2-48c2-ad21-59cd8d23f5b9.png)
7. Review the instructions. The script changes calendar events in a period of time (days before / after it runs). Enter how many days on the appropriate field
8. Enter the string you want to match in the title. You can enter several strings comma separated, no spaces before or after comma.
9. Now to run the script manually: SCRIPT_STARTER > Run Manually
10. Now to run the script automatically every 5 minutes: SCRIPT_STARTER > Run Automatically every 5 minutes. It will run even if spreadsheet is closed, computer turned off, etc
11. If you want to cancel automatic run SCRIPT_STARTER > Cancel Automatic Run

Best to test first to understand how it works:
a. Enter "Gobbledegook" on the spreadsheet under PALE_RED for example (C14)
b. Create a calendar event with title "The word of the year is Gobbledegook for 2022" for today
c. Run script manually
