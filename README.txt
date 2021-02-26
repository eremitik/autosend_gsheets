This is a repo for a simple script to download data from a google sheet, save into an excel file, and send the file automatically from your Gmail.

Dependicies:
pip install gspread
pip install pandas

Steps to setup:
1) Get your authorization credentials for Google services. Steps here: https://developers.google.com/identity/protocols/oauth2/web-server
2) Input the credentials into creds_pub.json file
3) Fill in the capitalised variables in send_email_pub.py

The script is setup to grab data from a specific sheet and cell range.
