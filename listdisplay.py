#!/usr/bin/env python
import tkinter as tk

def whichSelected():
    print("At "+str(select.curselection())+" of "+str(len(listB)))
    return int(select.curselection()[0])

def deleteEntry():
    del(listB[whichSelected()])
    setSelect()
def show():
    print(listB)

def setSelect():
    listB.sort()
    select.delete(0,tk.END)
    for name in listB:
        select.insert(tk.END,name)

listB=[1,2,3,4,5,6,7]
def Application():
    global ind
    global select

    win=tk.Tk()
    
    def __init__(self,master=None):
        self.listA=[]
        self.listB=[1,2,3,4,5,6,7]
        tk.Frame.__init__(self,master)
        self.grid(rowspan=20,columnspan=15)
        self.createWidget()

    frame = tk.Frame()
    frame.pack()

    b1 = tk.Button(frame,text="delete",command=deleteEntry)
    b1.pack(side=tk.LEFT)

    b2 = tk.Button(frame,text="show",command=show)
    b2.pack(side=tk.LEFT)
    
    scroll = tk.Scrollbar(frame,orient=tk.VERTICAL)
    select = tk.Listbox(frame,yscrollcommand=scroll.set)

    scroll.config(command=select.yview)
    scroll.pack(side=tk.RIGHT)

    select.pack(side=tk.LEFT,expand=1)
    return win

win = Application()
setSelect()
win.mainloop()
