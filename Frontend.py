import eLead
import mmcr
from tkinter import *
from tkinter.ttk import *
import mmcr2

class gui:

    def __init__(self):

        self.__window = Tk()
        self.__iconPhoto = PhotoImage(file="logo.png")

        self.__window.title("Service Drive Report Automation")
        self.__window.iconphoto(False,self.__iconPhoto)
        self.__window.config(bg = "gray17")
        self.__window.geometry("280x200")

        Label(self.__window,text="",background="gray17").grid(row=0)
        Label(self.__window,text="",background="gray17").grid(row=1)

        self.__dropDownLabel = Label(self.__window,text="Mode",background="gray17",foreground="white")
        self.__dropDownLabel.grid(row=2,column=0)
        self.__dropDown = Combobox(state="readonly",values=["Eleads","MMCR 1.0","MMCR 2.0"])
        self.__dropDown.grid(row=2,column=1)
        Label(self.__window,text="",background="gray17").grid(row=3)

        self.__emailLabel = Label(self.__window,text = "Email/Username",background="gray17",foreground="white")
        self.__emailLabel.grid(row=4,column=0)
        self.__emailEntry = Entry(self.__window)
        self.__emailEntry.grid(row=4,column=1)
        Label(self.__window,text="",background="gray17").grid(row=5)

        self.__passwordLabel = Label(self.__window,text = "Password",background="gray17",foreground="white")
        self.__passwordLabel.grid(row=6,column=0)
        self.__passswordEntry = Entry(self.__window)
        self.__passswordEntry.grid(row=6,column=1)
        Label(self.__window,text="",background="gray17").grid(row=7)

        self.__submitButton = Button(self.__window, text="Run",command=self.__buttonFunction)
        self.__submitButton.grid(row=8,column=1)

        self.__window.mainloop()
    
    def __buttonFunction(self):
        
        mode = self.__dropDown.get()
        user = self.__emailEntry.get()
        password = self.__passswordEntry.get()

        if mode == "Eleads":
            
            eLead.elead(user,password)
            #except:
            #    print("ERROR: Please Try Again")
        
        if mode == "MMCR 1.0":
            try:
                mmcr.checkServices(user=user,password=password)
            except Exception as e:
                print(e)
                print("ERROR: Please Try Again")

        if mode == "MMCR 2.0":
            try:
                mmcr2.checkServices(user=user,password=password)
            except:
                print("ERROR: Please Try Again")