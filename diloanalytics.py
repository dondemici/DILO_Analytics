#Created by dadriatico
#05Apr2021

import pandas as pd
import numpy as np

import os
from datetime import date
from datetime import datetime
from datetime import timedelta
import datetime as dt
import dateparser

import sys 
from tkinter import * 
import tkinter as tk
from tkinter import Button
from tkinter import Label
from tkinter import ttk
from tkinter import messagebox as mb
from tkcalendar import Calendar, DateEntry

def generateanalytics():

    #1 Import Data
    startTime = datetime.now()

    SDate = cal.get_date()
    SDate = pd.to_datetime(SDate)
    print(SDate)
    EDate = cal2.get_date()
    EDate = pd.to_datetime(EDate)
    print(EDate)

    dir_name=(e2.get())
    inputfile = (e1.get())
    fileloc0 = os.path.join(dir_name, inputfile) 
    df = pd.read_excel (fileloc0, sheet_name='DILO_Daily', usecols="A:N", engine='pyxlsb')
    df.dropna(subset = ["Date"], inplace=True)

    #1.2 Import Data 2 - Workshift
    df2 = pd.read_excel (fileloc0, sheet_name='Workshift', usecols="A:H", engine='pyxlsb')

    #Clean Time Data of Data 1
    df['Date'] = pd.TimedeltaIndex(df['Date'], unit='d') + dt.datetime(1899, 12, 30)
    df['Month'] = df['Date'].dt.to_period('m') 
    df['Week'] = df['Date'].dt.isocalendar().week 
    df['Start'] = (df['Start'] * 24)
    df['Start'] = np.round(df['Start'],2)
    df['End'] = (df['End'] * 24)
    df['Process Duration'] = df['End'] - df['Start']
    df['Process Duration'] = df['Process Duration'].round(2)
    df['Turnaround Time'] = df['Process Duration']/df['Volume Completed']

    #Combine Date and Time
    df["Date2"] = df["Date"].astype(str)
    df["Date2"] = df['Date2'].str[0:10]
    df['Start1'] = np.floor(df['Start']).astype(int)
    df['Start2'] = (df['Start']%(np.floor(df['Start'])))*60   
    df['Start2'] = df['Start2'].astype(int)
    df['Start2']=df['Start2'].apply(lambda i: '{0:0>2}'.format(i))

    df['End1'] = np.floor(df['End']).astype(int)
    df['End2'] = (df['End']%(np.floor(df['End'])))*60   
    df['End2'] = df['End2'].astype(int)
    df['End2']=df['End2'].apply(lambda i: '{0:0>2}'.format(i))

    df['Start Time'] = df["Date2"] + ' ' + df["Start1"].astype(str) + ':' + df["Start2"].astype(str) 
    df['Start Time'] = pd.to_datetime(df['Start Time'])
    df['End Time'] = df["Date2"] + ' ' + df["End1"].astype(str) + ':' + df["End2"].astype(str) 
    df['End Time'] = pd.to_datetime(df['End Time'])
    df['Turnaround Time'] = df['Turnaround Time'].round(2)
    
    #To be used as filename
    username = df.iloc[0] ["Process Owner"]

    #Daily FTE Sheet
    dfvol = (df.pivot_table(index=['Process Owner','Activities'], columns='Date', values='Process Duration',
                aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
    .sort_values('All', ascending=False))  # sort by row sum

    #Summarized DILO Sheet
    dfvol2 = df.pivot_table( 
        values=['Process Duration','Volume Completed'],
        index=['Month','Week','Process Owner','Date','Activities'], 
        aggfunc=np.sum
    ).reset_index() 
    dfvol2 = (dfvol2.loc[:100000,["Month","Week","Process Owner","Date","Activities","Process Duration","Volume Completed"]]) 
    dfvol2['Process Duration'] = dfvol2['Process Duration'].round(2)
    dfvol2['Turnaround Time'] = dfvol2['Process Duration']/dfvol2['Volume Completed']
    dfvol2['Turnaround Time'] = dfvol2['Turnaround Time']*60
    dfvol2['Turnaround Time'] = dfvol2['Turnaround Time'].round(2)


    #Average FTE
    dfvol3 = (dfvol2.pivot_table(index=['Process Owner','Activities'], columns='Week', values='Process Duration',
                aggfunc='median', fill_value=0, margins=True)   # pivot with margins 
    .drop('All')
    .sort_values('All', ascending=False))
    dfvol3 = np.round(dfvol3, 2)

    #Volume
    dfvol4 = (dfvol2.pivot_table(index=['Process Owner','Activities'], columns='Week', values='Volume Completed',
                aggfunc=['sum','mean'], fill_value=0, margins=True))   # pivot with margins 
    dfvol4 = np.round(dfvol4, 2)

    #Turnaround Time
    dfvol5 = (dfvol2.pivot_table(index=['Process Owner','Activities'], columns='Week', values='Turnaround Time',
                aggfunc=['mean'], fill_value=0, margins=True))   # pivot with margins 
    dfvol5 = np.round(dfvol5, 2)

    #WORKSHIFT DATA
    #Clean Time Data of Workshift Data Frame
    df2['Effectivity_Start'] = pd.TimedeltaIndex(df2['Effectivity_Start'], unit='d') + dt.datetime(1899, 12, 30)
    df2['Effectivity_End'] = pd.TimedeltaIndex(df2['Effectivity_End'], unit='d') + dt.datetime(1899, 12, 30)
    df2['Workshift_In'] = (df2['Workshift_In'] * 24)
    df2['Workshift_In'] = np.round(df2['Workshift_In'],2)
    df2['Workshift_Out'] = (df2['Workshift_Out'] * 24)
    df2['Workshift_Out'] = np.round(df2['Workshift_Out'],2)
    df2.dropna(subset = ["Effectivity_Start"], inplace=True)

    df3 = pd.concat([pd.DataFrame({'Effectivity_Start': pd.date_range(row.Effectivity_Start, row.Effectivity_End, freq='d'),
                'Workshift_In': row.Workshift_In, 'Last_Name': row.Last_Name,'First_Name': row.First_Name,
                'Workshift_Out': row.Workshift_Out}, columns=['Last_Name','First_Name','Effectivity_Start','Workshift_In','Workshift_Out'])
        for i, row in df2.iterrows()], ignore_index=True)
    
    df3["Date2"] = df3["Effectivity_Start"].astype(str)
    df3["Date2"] = df3['Date2'].str[0:10]

    df3['Workshift Start'] = df3["Date2"] + ' ' + df3["Workshift_In"].astype(str) + ':00' 
    df3['Workshift Start'] = pd.to_datetime(df3['Workshift Start'])
    df3['Workshift End'] = df3["Date2"] + ' ' + df3["Workshift_Out"].astype(str) + ':00' 
    df3['Workshift End'] = pd.to_datetime(df3['Workshift End'])

    #df3 = (df3.loc[:10000,['Last_Name', 'First_Name', 'Effectivity_Start','Workshift Start', 'Workshift End']])  

    #TIMESHEETS DATA
    try:
        start_date = SDate #"3/21/2020"
        print(start_date)
        end_date = EDate #"12/31/2022"
        print(end_date)
        after_start_date = df["Start Time"] >= start_date
        before_end_date = df["End Time"] <= end_date
        between_two_dates = after_start_date & before_end_date
        df = df.loc[between_two_dates]

        dfs = df.sort_values(['Start Time'], ascending=True)
        dfs = df.drop_duplicates(subset= ["Date"], keep='first')
        dfs = (dfs.loc[:10000,["Process Owner","Date","Start Time"]]) 

        dfe = df.sort_values(['End Time'], ascending=True)
        dfe = df.drop_duplicates(subset= ["Date"], keep='last')
        dfe = (dfe.loc[:10000,["Process Owner","Date","End Time"]]) 

        dfs.reset_index(drop=True, inplace=True)
        dfe.reset_index(drop=True, inplace=True)
        df_total = pd.concat( [dfs, dfe], axis=1) 
        #df_total = (df_total[:10000,["Name",
        #    "Start Time","End Time"]]) 

        df_total = (df_total.iloc[:10000,[0,1,2,5]]) 
        df_total["Date3"] = df_total["Date"].astype(str)
        df_total['Workshift In Calc'] = df_total["Date3"] + ' ' + "6:15 AM"
        df_total['Workshift In'] = df_total["Date3"] + ' ' + "6:00 AM"
        df_total['Workshift Out'] = df_total["Date3"] + ' ' + "3:00 PM"
        df_total['Workshift In Calc'] = pd.to_datetime(df_total['Workshift In Calc'])
        df_total['Workshift In'] = pd.to_datetime(df_total['Workshift In'])
        df_total['Workshift Out'] = pd.to_datetime(df_total['Workshift Out'])

        #Overwrite Workshift
        df3.rename(columns={'Effectivity_Start':'Date'}, inplace=True)
        df4 = pd.merge(df_total, df3[['Date','Workshift Start','Workshift End','Workshift_Out']], on='Date', how='left')
        #df4.loc[(df4['Workshift Start'] == np.NaN), 'Workshift Start'] = df4['Workshift In']
        df4.loc[df4["Workshift Start"].isnull(),'Workshift Start'] = df4['Workshift In']
        df4.loc[df4["Workshift End"].isnull(),'Workshift End'] = df4['Workshift Out']
        df4["Workshift Start Calc"] = df4["Workshift Start"] + timedelta(minutes=15)
        print(df4)

        #Define Tardiness
        Tardiness = df4['Start Time'] - df4['Workshift Start Calc']
        df4['Tardiness Calc'] = (Tardiness / np.timedelta64(1, 'h'))
        df4['Tardiness Calc'] = df4['Tardiness Calc'].round(2)
        df4.loc[(df4['Tardiness Calc'] <= 0), 'Tardiness'] = 0
        df4.loc[(df4['Tardiness Calc'] > 0), 'Tardiness'] = df4['Tardiness Calc']

        #Define OT
        OT = df4['End Time'] - df4['Workshift End']
        df4['Overtime'] = (OT / np.timedelta64(1, 'h'))
        df4['Overtime'] = df4['Overtime'].round(2)

        #Define Remarks
        df4.loc[(df4['Tardiness Calc'] >= 4), 'System Remarks'] = "Confirm if Half Day"
        df4.loc[(df4['Overtime'] > 0), 'System Remarks'] = "Overtime"
        df4.loc[(df4['Overtime'] >= 4), 'System Remarks'] = "Confirm if Half Day"
        df4.loc[(df4['Overtime'] < 0), 'System Remarks'] = "Undertime"
        df4.loc[(df4['Overtime'] < -5), 'System Remarks'] = "Time Out may be Incorrect"
        df4['Manual Remarks'] = df['OT Reason']
        print(df['OT Reason'])
        df4 = (df4.iloc[:10000,[0,1,2,3,8,9,13,14,15,16]]) 
        df4a = (df4.iloc[:10000,[0,2,3,4,5,6,7,8,9]]) 

        df5 = pd.DataFrame()
        df5['Time Entries'] = df4['Start Time'].append(df4['End Time']).reset_index(drop=True)
        lastnam = df.iloc[0] ["Process Owner"]
        df5['Process Owner'] = lastnam
        df5['Manual Remarks'] = df4['Manual Remarks']
        df5 = df5.sort_values(['Time Entries'], ascending=True)
        df5 = (df5.iloc[:10000,[1,0,2]]) 
    except ValueError:
        pass

    #OT Calculation Data
    df3.rename(columns={'Effectivity_Start':'Date'}, inplace=True)
    df6 = pd.merge(df, df3[['Date','Workshift Start','Workshift End','Workshift_Out']], on='Date', how='left')
    df6['Workshift_Out'].fillna(15, inplace=True)
    df6['End2'] = df6['End2'].astype(int)
    df6['Time Results'] = df6['Workshift_Out'] - df6['End1']
    df6['Time Results2'] = df6['Workshift_Out'] - df6['Start1']
    conditions = [
        (df6['Time Results'] > 0),
        (df6['Time Results'] == 0) & (df6['End2'] == 0),
        (df6['Time Results'] == 0) & (df6['End2'] > 0),
        (df6['Time Results'] < 0),
    ]
    results =['Retain','Retain','Remove', 'Remove']
    df6['Data_Filter'] = np.select(conditions,results)
    
    conditions2 = [
        (df6['Data_Filter'] == 'Remove') & (df6['Time Results'] == 0),
        (df6['Data_Filter'] == 'Remove') & (df6['Time Results'] < 0),
        (df6['Data_Filter'] == 'Retain'),
    ]
    results2 =['Update','Remove','Retain',]
    df6['Data_Filter2'] = np.select(conditions2,results2)
    
    #Generate Output

    print(username)

    outputfile = ('DILO Analysis ' + username + '.xlsx') 
    outfileloc = os.path.join(dir_name, outputfile) 
    writer = pd.ExcelWriter(outfileloc) 

    try:
        df5.to_excel(writer, sheet_name = 'TS for JP Upload') 
        df4.to_excel(writer, sheet_name = 'Timesheets') 
    except UnboundLocalError:
        pass
    df6.to_excel(writer, sheet_name = 'OT Calc')
    dfvol.to_excel(writer, sheet_name = 'Daily FTE')
    dfvol3.to_excel(writer, sheet_name = 'Average FTE per Week')
    dfvol4.to_excel(writer, sheet_name = 'Volume')
    dfvol5.to_excel(writer, sheet_name = 'Turnaround Time')
    dfvol2.to_excel(writer, sheet_name = 'Summarized DILO Data') 
    df.to_excel(writer, sheet_name = 'Raw DILO Data')
    df3.to_excel(writer, sheet_name = 'Workshift Data')
    writer.save()

    print(datetime.now() - startTime)     
    print("Analytics generated!")
    
    try:
        dframe = df4a
        txt = Text(tab2,width=160, height=30) 
        txt.pack()
        #txt.grid(row=0, column=1, sticky=tk.SW, pady=20,padx=40, columnspan=10) 
        class PrintToTXT(object): 
            def write(self, s): 
                txt.insert(END, s, str(df4.iloc[:6,1:2]))
        sys.stdout = PrintToTXT() 
        print ('Timesheets') 
        print (dframe)
    except UnboundLocalError:
        pass

    mb.showinfo('Message', 'Analytics generated!')

def show_entry_fields():
    print("Path: %s" % (e2.get()))
    mb.showinfo('Your DILO File Folder', "Path: %s" % (e2.get()))

master = tk.Tk()
master.title("DILO Analytics Generator")
master.geometry("700x450")

tabControl = ttk.Notebook(master)
  
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
  
tabControl.add(tab1, text ='DILO Analytics')
tabControl.add(tab2, text ='Timesheets')
tabControl.pack(expand = 1, fill ="both")


# Add Calender
today = dt.date.today()
ttk.Label(tab1, text='Start Date').grid(row=0, column=1, sticky=tk.SW, pady=20,padx=40)
cal = DateEntry(tab1, width=12, background='darkblue',
                    foreground='white', borderwidth=2, year=today.year, month=today.month, day=today.day - 15)
cal.grid(row=0, column=1, pady=20,padx=40)
SDate = cal.get_date()
print (SDate)

ttk.Label(tab1, text='End Date').grid(row=0, column=2, sticky=tk.SW, pady=20)
cal2 = DateEntry(tab1, width=12, background='darkblue',
                    foreground='white', borderwidth=2, year=today.year, month=today.month, day=today.day)
cal2.grid(row=0, column=2, sticky=tk.E, pady=20,padx=40)
EDate = cal2.get_date()
print (EDate)


tk.Label(tab1,
    text="DILO Input File Name").grid(row=3, column=1, sticky=tk.W, padx=40)

e1 = tk.Entry(tab1, width=80)
e1.grid(row=3, column=1, sticky=tk.E, padx=170, columnspan=8)

tk.Label(tab1,
    text="DILO Input File Path").grid(row=4, column=1, sticky=tk.W, padx=40)

e2 = tk.Entry(tab1, width=80)
e2.grid(row=4, column=1, sticky=tk.E, padx=170, pady=5, columnspan=8)

tk.Button(tab1, text='Generate Analytics', command=generateanalytics).grid(row=5, column=1, sticky=tk.W, pady=10, padx=40)
tk.Button(tab1, text='Show Full Path', command=show_entry_fields).grid(row=5, column=1, sticky=tk.E, padx=120, pady=10)
tk.Button(tab1, text='Quit', command=tab1.quit).grid(row=5, column=2, sticky=tk.W, pady=10)


tk.mainloop()



