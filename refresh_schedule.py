import schedule
import time
import win32com.client

def daily_bi_queries():
# Open Excel
    path='C:\\Users\\Michael\\Downloads\\'
    FileName='source_file_Austria.xlsx'
    Application = win32com.client.Dispatch("Excel.Application")
 
 # Show Excel. While this is not required, it can help with debugging
    Application.Visible = 1
    Application.DisplayAlerts=False
    Application.AskToUpdateLinks = False
 # Open Your Workbook
    Workbook = Application.Workbooks.open(path + FileName)
    #try:
    #    Workbook.UpdateLink(Name=Workbook.LinkSources())

   # except Exception as e:
    #    print(e)
    # Refesh All
    Workbook.RefreshAll()
    print('Done')
 # Saves the Workbook
    Workbook.Save()
 
 # Closes Excel
    Application.Quit()


#schedule.every(1).minutes.do(daily_bi_queries)
#schedule.every().hour.do(job)
#schedule.every().day.at("10:30").do(daily_bi_queries)
#schedule.every().monday.do(job)
#schedule.every().wednesday.at("13:15").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)