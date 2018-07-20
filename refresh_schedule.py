import schedule
import time
import win32com.client

def refresh_files(file_list,time_now):
    for file in file_list:

        # Start an instance of Excel
        xlapp = win32com.client.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
        wb = xlapp.workbooks.open(file)

    # Optional, e.g. if you want to debug
    # xlapp.Visible = True

    # Refresh all data connections.
        wb.Model.Refresh()
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close(True)

        print("All is done")
        # Quit
        xlapp.Quit()
        print("I'm working...")





#schedule.every(10).minutes.do(refresh_files)
#schedule.every().hour.do(job)
#schedule.every().day.at("10:30").do(job)
#schedule.every(5).to(10).minutes.do(job)
#schedule.every().monday.do(job)
#schedule.every().wednesday.at("13:15").do(job)

#while True:
    #schedule.run_pending()
cur_time = time.localtime(time.time())
time_now=str(cur_time[3]) + ':' + str(cur_time[4])

