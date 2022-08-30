from tkinter import *
from tkinter import ttk
import re
from tkcalendar import Calendar
import time
from tkinter import messagebox
from datetime import timedelta
import tkcalendar
from EGM import EGM_account,EGM_Players
from EGM_Machine_Data_Inte import EGM_inte_machinedata
from EGM_Players_Data_inte import EGM_inte_Playerdata



#validation
def validation():
    EGM=CheckVar2.get()
    if EGM==1:
        msg = 'EGM data is processing..........'
        messagebox.showinfo('message', msg)
        start = time.time()
        msg = 'EGM Account data is processing..... '
        messagebox.showinfo('message', msg)
        bar(20)
        EGM_account(dates)
        msg = 'EGM Account data is done'
        messagebox.showinfo('message', msg)
        bar(40)
        msg = 'EGM Player data is processing......'
        messagebox.showinfo('message', msg)
        EGM_Players(dates)
        msg = 'EGM Account Player is done'
        messagebox.showinfo('message', msg)
        bar(60)
        msg = 'EGM data is integrating......'
        messagebox.showinfo('message', msg)
        msg = 'EGM floor analysis data is integrating......'
        messagebox.showinfo('message', msg)
        EGM_inte_machinedata()
        msg = 'EGM floor analysis data is done.'
        messagebox.showinfo('message', msg)
        bar(80)
        msg = 'EGM Loyal data is intergrating......'
        messagebox.showinfo('message', msg)
        EGM_inte_Playerdata()
        bar(100)
        msg = 'EGM all Completed'
        messagebox.showinfo('message', msg)
        # msg = 'EGM VLD data is done......'
        # messagebox.showinfo('message', msg)
        end = time.time()
        lengh=end-start
        msg = "{} {} {}".format("All done, spend:", lengh/60,"Mins")
        messagebox.showinfo('message', msg)












#input the week name

window = Tk()

window.title("Welcome to Data Collection")

progress = ttk.Progressbar(window, orient = HORIZONTAL,length=200, mode='determinate')
# progress.pack(pady = 10)
progress.place(x=100, y=400,height=30)

def update_progress_label():
    return f"Current Progress: {progress['value']}%"

value_label = ttk.Label(window,text=update_progress_label())
value_label.place(x=100, y=450,height=30)


def bar(x):
    progress['value']=x

    if progress['value']==20:
        window.update_idletasks()
        value_label['text'] = update_progress_label()

    if progress['value'] ==40:
        window.update_idletasks()
        value_label['text'] = update_progress_label()

    if progress['value'] ==60:
        window.update_idletasks()
        value_label['text'] = update_progress_label()

    if progress['value'] ==80:
        window.update_idletasks()
        value_label['text'] = update_progress_label()

    if progress['value'] ==90:
        window.update_idletasks()
        value_label['text'] = update_progress_label()
    if progress['value'] ==100:
        value_label['text'] = update_progress_label()
        msg = 'EGM Account data is done'
        messagebox.showinfo('message', msg)



window.geometry('1000x600')
string='Please choose the area'

CheckVar1 = IntVar()
CheckVar2 = IntVar()
CheckVar3 = IntVar()
CheckVar4 = IntVar()
C1 = Checkbutton(window, text = "Keno", variable = CheckVar1,background='#2E86C1',
                 onvalue = 1, offvalue = 0, height=20,font = ('calibre',25,'bold'),
                 width =5)
C2 = Checkbutton(window, text = "EGM", variable = CheckVar2,font = ('calibre',25,'bold'),background='#2E86C1',
                 onvalue =1, offvalue =0, height=20,
                 width =5)
C3 = Checkbutton(window, text = "ATG", variable = CheckVar3,background='#2E86C1',
                 onvalue = 1, offvalue = 0, height=20,font = ('calibre',25,'bold'),
                 width =5)
C4 = Checkbutton(window, text = "TG", variable = CheckVar4,font = ('calibre',25,'bold'),background='#2E86C1',
                 onvalue =1, offvalue =0, height=20,
                 width =5)

C1.place(x=30, y=250,height=30)
C2.place(x=200, y=250,height=30)
C3.place(x=50, y=300,height=30)
C4.place(x=210, y=300,height=30)

label_title = Label(window, text=string,font = ('calibre',35,'bold'),fg='#000',background='#2E86C1')
label_date = Label(window, text="Date:",font = ('calibre',25,'bold'),fg='#fff',background='#2E86C1')

label_title.place(x=250,y=50)
# label_date.place(x=30, y=400)

Date_txt_ = Label(window, text="" ,font = ('calibre',7,"bold"),width=80)
Date_txt_.place(x=400, y=450,height=100)
#date calendar

def date_range(start, stop):
    global dates # If you want to use this outside of functions

    # dates = []
    diff = (stop - start).days
    dates = [start + timedelta(days=i) for i in range(diff + 1)]
    # for i in range(diff + 1):
    #     day = start + timedelta(days=i)
    #     dates.append(day)
    dates = [x.strftime('%Y-%m-%d') for x in dates]
    if dates:
        return # Print it, or even make it global to access it outside this
    else:
        print('Make sure the end date is later than start date')

def grad_date():
    Date_txt_.config(text="Selected Date is: " + ' '.join(dates))


date1 = tkcalendar.DateEntry(window,date_pattern="dd/mm/yyyy")
date1.place(x=400, y=300,width=150)

date2 = tkcalendar.DateEntry(window,date_pattern="dd/mm/yyyy")
date2.place(x=400, y=350,width=150)

Button(window, text='Find range', command=lambda: [date_range(date1.get_date(), date2.get_date()),grad_date()]).place(x=400, y=400,width=100)

# Date_txt_.config(text="Selected Date is: " + dates)


# cal = Calendar(window, selectmode='day',
#                year=2022, month=2,
#                day=14,date_pattern="dd/mm/yyyy")
# cal.place(x=600, y=250,width=300)
#









#

# Date_txt = Entry(window,font = ('calibre',25,'bold'),width=16)
# Date_txt.place(x=350, y=400,height=50)

# cal_Button=Button(window, text="Get Date",font = ('calibre',10,'bold'),command=grad_date)
btn_refresh = Button(window, text="Avilability of Data",font = ('calibre',25,'bold'))
btn = Button(window, text="Submit",font = ('calibre',25,'bold'),command=validation)
btn2=Button(window, text="Quit",font = ('calibre',25,'bold'),command=window.destroy)

# cal_Button.place(x=700, y=550,width=100)
btn_refresh.place(x=350, y=150,width=300)
btn.place(x=30, y=500,width=130)
btn2.place(x=200, y=500,width=130)
window.configure(background='#2E86C1')


#center the windown on the screen

windowWidth = window.winfo_reqwidth()
windowHeight = window.winfo_reqheight()
# Gets both half the screen width/height and window width/height
positionRight = int(window.winfo_screenwidth() / 5 - windowWidth / 2)
positionDown = int(window.winfo_screenheight() / 5 - windowHeight / 2)

# Positions the window in the center of the page.
window.geometry("+{}+{}".format(positionRight, positionDown))
window.mainloop()
