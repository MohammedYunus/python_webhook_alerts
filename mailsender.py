import pandas as pd
import time
import win32com.client
from datetime import datetime, timedelta
from pytz import timezone
import os
import schedule
import requests
import json

cwd = os.getcwd()
intz = timezone('Asia/Kolkata')

err_url = "https://hooks.platform.aws/incomingwebhooks"

def err_desc(alrt):
    data = {'Content':alrt}
    requests.post(err_url, data=json.dumps(data), headers={'Content-Type':'application/json'})

def file_organize():
    try:
        now = datetime.now()
        yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        my_csv_data = pd.read_csv('log_file.csv')
        log_file_df = pd.DataFrame(my_csv_data)
        total_length = len(log_file_df.index)

        if (now.hour == 23 and now.minute == 59):
            log_file_df.to_csv(f'log_files_archive/log_file.csv {yesterday}', index=False)

        elif (now.hour == 0 and now.minute == 2):
            for delete_row in range(0, total_length):
                log_file_df = log_file_df.drop(index=[delete_row])
            log_file_df.to_csv('log_file.csv', index=False)

    except Exception as e:
        dateSt = datetime.today().date()
        my_csv_data = pd.read_csv('log_file.csv')
        log_file_df = pd.DataFrame(my_csv_data)
        total_length = len(log_file_df.index)
        log_file_df.to_csv(f'log_files_archive/log_file.csv {dateSt}', index=False)
        print(e)

def send_file(batch, today, filename, mailto, mailCC):
    try:
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = mailto
        mail.CC = mailCC
        mail.Subject = f'{filename} {today} {batch}'

        body = f'''
            <html>
            <head>
            <style>body{{font-family:Calibri, sans-serif;}}
            </head>
            <body>
            <p>Hi Team,<br><br>
            Please find the attached report.
            </p>
            Regards,<br>
            Mohammed Yunus
            </body>
            </html>
        '''

        mail.HTMLBody = body
        base_in_ref_file_path = rf'{cwd}\log_file.csv'
        if base_in_ref_file_path:
            mail.Attachments.Add(base_in_ref_file_path)

        mail.display()
        mail.Send()
        print('mail sent')
        err_desc(f"Mail sent successfully on the {batch} batch time.")

    except Exception as e:
        print('mail not sent: ', e)
        err_desc("Mail not sent. batch preparing failed...")

def auto_prepare():
    try:
        now = datetime.now()

        today = datetime.today().strftime("%Y-%m-%d")
        print(today)

        # yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        # print(yesterday)

        if (now.hour == 0):
            batch = '12_00_AM'

        elif (now.hour == 4):
            batch = '04_00_AM'

        elif (now.hour == 8):
            batch = '08_00_AM'

        elif (now.hour == 12):
            batch = '12_00_PM'
        
        elif (now.hour == 16):
            batch = '04_00_PM'

        elif (now.hour == 20):
            batch = '08_00_PM'

        else:
            print('incorrect batch time:', now)
            
        filename = 'Sample Report'

        mailto = 'mohammedyunus@xyz.com' 
        mailCC = 'mohammedyunus@xyz.com'

        current_date = datetime.today().strftime("%d")
        print("-------------------------*date="+current_date+"*-------------------------")
        send_file(batch, today, filename, mailto, mailCC)
        
    except Exception as e:
        print('batch preparing failed: ',e)


schedule.every().day.at("00:00").do(auto_prepare)
schedule.every().day.at("04:00").do(auto_prepare)
schedule.every().day.at("08:00").do(auto_prepare)
schedule.every().day.at("12:00").do(auto_prepare)
schedule.every().day.at("16:00").do(auto_prepare)
schedule.every().day.at("20:00").do(auto_prepare)

print('First run: ',datetime.now().strftime("%d-%m-%Y %H:%M:%S"))

while True:
    schedule.run_pending()
    file_organize()
    time.sleep(1) 
