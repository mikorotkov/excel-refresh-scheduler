import schedule
import time
import win32com.client

def refresh_files(file_list,time_now):
    error_list=[]
    for file in file_list:

        # Start an instance of Excel
        print(file)
        xlapp = win32com.client.DispatchEx("Excel.Application")

        # Open the workbook in said instance of Excel
        wb = xlapp.workbooks.open(file)

        # Optional, e.g. if you want to debug
        xlapp.Visible = True
        print('Working on file {}'.format(file))

        # Refresh all data connections.
        wb.Model.Refresh()
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        try:
            wb.Save()
            print('Saved')
        except: 
            error_list.append(file)
            print('File was not saved')
            continue
        wb.Close(True)
            

    # Quit
    xlapp.Quit()






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
refresh_files(['Y:\\Draft - CA dashboard v2.xlsx'],time_now)

