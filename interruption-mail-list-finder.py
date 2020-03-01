import pyodbc
from tkinter import *
import tkinter


def data_getir():
    cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"  #connection sql
                          "Server=10.236.51.40,1433;"
                          "Database=;"
                          "uid=")
    response = cnxn.cursor();
    response.execute("select * from Customers") #find all customers
    column = [column[0] for column in response.description]
    data = []
    for row in response.fetchall():
        data.append(dict(zip(column, row)))
    return data




def select_all():
    lb.select_set(0,END)



def select_all2():
    root.clipboard_clear()
    a=(lb3.get(0,END))
    a=''.join(a)
    root.clipboard_append(a)
    cliptext= root.clipboard_get()
    print(cliptext)


def secimi_kaldir():
    lb.selection_clear(0, END)



def select_all3():
        lb3.delete(0,END)
        lb2.delete(0,END)
        cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                              "Server=;"
                              "Database=;"
                              "uid=;pwd=")


        my_tuple = lb.curselection()
        my_list = []
        for item in my_tuple:
            x = lb.get(item)
            y = x.split("-")
            my_list.append(y[0])

        temp_str = "( "
        for item in my_list:
            temp_str += item
            if item != my_list[-1]:
                temp_str += " , "
        temp_str += ")"
        try:
            response = cnxn.cursor();
            response.execute("select * from Customers where ID in " + temp_str)
            column = [column[0] for column in response.description]

            data = []
            for row in response.fetchall():
                data.append(dict(zip(column, row)))
            for item in data:
                if item["ContactP1"] != None:
                    lb2.insert(END, "{}".format(item["ContactP1"]).rstrip()+"," )
                else:
                    pass
                if item["ContactP2"] != None:
                    lb2.insert(END, "{}".format(item["ContactP2"]).rstrip() + ",")
                else:
                    pass
                if item["ContactP3"] != None:
                    lb2.insert(END, "{}".format(item["ContactP3"]).rstrip() + ",")
                else:
                    pass
                if item["Email1"]!= None:
                    lb3.insert(END, "{}".format(item["Email1"]).rstrip()+ ";")
                else:
                    pass
                if item["Email2"]!= None:
                    lb3.insert(END, "{}".format(item["Email2"]).rstrip() + ";")
                else:
                    pass
                if item["Email3"]!= None:
                    lb3.insert(END, "{}".format(item["Email3"]).rstrip() + ";")
                else:
                    pass
        except:
            pass

def btn2_click(evt):
        lb3.delete(0,END)
        lb2.delete(0,END)
        cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                              "Server=;"
                              "Database=;"
                              "uid=;pwd=")


        w = evt.widget
        my_tuple = w.curselection()

        for tuple in my_tuple:
            w.get(tuple)


        my_list = []
        for item in my_tuple:
            x = lb.get(item)
            y = x.split("-")
            my_list.append(y[0])

        temp_str = "( "
        for item in my_list:
            temp_str += item
            if item != my_list[-1]:
                temp_str += " , "
        temp_str += ")"
        try:
            response = cnxn.cursor();
            response.execute("select * from Customers where ID in " + temp_str)
            column = [column[0] for column in response.description]

            data = []
            for row in response.fetchall():
                data.append(dict(zip(column, row)))
            for item in data:
                if item["ContactP1"] != None:
                    lb2.insert(END, "{}".format(item["ContactP1"]).rstrip()+"," )
                else:
                    pass
                if item["ContactP2"] != None:
                    lb2.insert(END, "{}".format(item["ContactP2"]).rstrip() + ",")
                else:
                    pass
                if item["ContactP3"] != None:
                    lb2.insert(END, "{}".format(item["ContactP3"]).rstrip() + ",")
                else:
                    pass
                if item["Email1"]!= None:
                    lb3.insert(END, "{}".format(item["Email1"]).rstrip()+ ";")
                else:
                    pass
                if item["Email2"]!= None:
                    lb3.insert(END, "{}".format(item["Email2"]).rstrip() + ";")
                else:
                    pass
                if item["Email3"]!= None:
                    lb3.insert(END, "{}".format(item["Email3"]).rstrip() + ";")
                else:
                    pass
        except:
            pass


root = Tk()
root.geometry('600x500')
root.title("Kesinti Bildirimi")
B = tkinter.Button(root, text="Seçimi Kaldır", command=secimi_kaldir)
B.pack()
B.place(x=450, y=50)

select = tkinter.Button(root, text="Hepsini Seç", command=select_all)
select.pack()
select.place(x=450, y=75)

label = Label(root, text="Müşteri Grubu").place(x=0, y=0)
frame = Frame(root)
frame2 = Frame(root)















frame.place(x=5, y=20)  # Position of where you would place your listbox
lb = Listbox(frame, width=70, height=5, selectmode=MULTIPLE)
lb.bind('<<ListboxSelect>>', btn2_click)
lb.pack(side='left', fill='y')
scrollbar = Scrollbar(frame, orient="vertical", command=lb.yview)
scrollbar.pack(side="right", fill="y")
lb.config(yscrollcommand=scrollbar.set)


lb.delete(0, END)

customerdata = data_getir()
for customer in customerdata:
    lb.insert(END, "{}-{}".format(customer["ID"], customer["Customer"])),
lb.delete(0)
lb.delete(37)


label2 = Label(root, text="Yetkili Kişi İsim").place(x=0, y=150)
frame = Frame(root)
frame.place(x=5, y=170)  # Position of where you would place your listbox
lb2 = Listbox(frame, width=70, height=5)
lb2.pack(side='left', fill='y')
scrollbar = Scrollbar(frame, orient="vertical", command=lb2.yview)
scrollbar.pack(side="right", fill="y")
lb2.config(yscrollcommand=scrollbar.set)
B2 = tkinter.Button(root, text="Tümünü Getir/Temizle", command=select_all3)
B2.pack()
B2.place(x=450, y=200)


label3 = Label(root, text="Yetkili Kişi Mail").place(x=0, y=300)
frame = Frame(root)
frame.place(x=5, y=330)  # Position of where you would place your listbox
lb3 = Listbox(frame, width=70, height=5)
lb3.pack(side='left', fill='y')
scrollbar = Scrollbar(frame, orient="vertical", command=lb3.yview)
scrollbar.pack(side="right", fill="y")
lb3.config(yscrollcommand=scrollbar.set)

B3 = tkinter.Button(root, text="Mailleri Seç/cmKopyala", command=select_all2)
B3.pack()
B3.place(x=450, y=360)






root.mainloop()

