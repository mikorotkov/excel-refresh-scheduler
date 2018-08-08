import schedule
import time
import win32com.client
#import shutil

def daily_bi_queries():
# Open Excel
 #   Application = win32com.client.Dispatch("Excel.Application")
 
# Show Excel. While this is not required, it can help with debugging
 #   Application.Visible = 1
 
# Open Your Workbook
 #   Workbook = Application.Workbooks.open(SourcePathName + '/' + FileName)
 
# Refesh All
  #  Workbook.RefreshAll()
 
# Saves the Workbook
 #   Workbook.Save()
 
# Closes Excel
 #   Application.Quit()
    print("I'm working...")
    print(f'SourcePathName is {SourcePathName} and FileName is {FileName}')

schedule.every(1).minutes.do(daily_bi_queries())
#schedule.every().hour.do(job)
#schedule.every().day.at("10:30").do(job)
#schedule.every().monday.do(job)
#schedule.every().wednesday.at("13:15").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)