import eLead
import mmcr
from tkinter import *
from tkinter.ttk import *

class gui:


    def __init__(self):

        self.__netstarUsername = None
        self.__netstarPassword = None
        self.__eLeadsEmail = None
        self.__eLeadsPassword = None
        self.__window = Tk()
        self.__iconPhoto = PhotoImage(file="logo.png")


        self.__window.title("Service Drive Report Automation")
        self.__window.iconphoto(False,self.__iconPhoto)
        self.__window.config(bg = "gray17")
        self.__window.geometry("500x400")


        self.__window.mainloop()

gui()
