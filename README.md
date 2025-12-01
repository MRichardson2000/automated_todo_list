💻 Automated Todo List

This script automates your todo list by reading your outlook calendar and then it emails you your meetings for the day. I build this
because all of my work was being scheduled in my calendar so instead of it being written in a notepad, it emails it to me every morning. 
Then I also get it to email me once a week with my entire week so I have a full overview of the week ahead. It helped with my productivity and management. 

-- THIS IS ON MY LAPTOP CURRENTLY. IT WILL NOT WORK ON A VM --

📦 Features

Scans your outlook calendar
Identifies all meetings within the time frame
Sends a summary email via Outlook to the email specified

🧰 Requirements

Python 3.8+
pywin32 (for Outlook integration)

Install dependencies: run uv sync to pull dependencies from the toml file

📁 File Structure project/ 
│
 ├── main1.py # Main script for 1 day reminders
 ├── main7.py # Main script for 7 day reminders
 └── README.md # Documentation

🚀 Usage Set up a batch file with the below, put your user path in and then specify on your c drive where you cloned the repo. This is just where most of my stuff went: C:\Users\YOURUSERPATHHERE\AppData\Local\Programs\Python\Python313\python.exe C:\Utilities\Python\eol_laptop_reminders\main.py Then set up a schedule task to run once a day/week on your chosen day and time that targets this batch file and runs it.

This will:

Read your outlook calendar
Check for meetings 
Send an email to the email address with the list of upcomingtasks for the day/week

✉️ Email Setup The script uses win32com.client to send emails via Outlook. The recipient is currently set to: mail.To = ""

You can change this to any valid email address or distribution list.

📌 Notes

run a uv sync in the terminal to pull the dependencies from the toml
if you need uv, run pip install uv in the terminal to add to your global scope, then run uv sync