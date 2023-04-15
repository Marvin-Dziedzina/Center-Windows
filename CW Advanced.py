from pygetwindow import getAllWindows, getWindowsWithTitle
from tkinter import Tk
from tkinter import Label, Button, BooleanVar, Checkbutton
from keyboard import add_hotkey
from webbrowser import open as wbOpen
from shutil import copy2
from win32com.client import Dispatch as win32comDiscpatch
from os import getlogin, environ, remove
from os.path import dirname, realpath, isfile, join
from time import time, sleep

# Version with in this format
# X.Y.Z
# X = Major Version
# Y = Minor Version
# Z = Patch
#
VERSION = "1.2.1"

programPath = dirname(realpath(__file__))
picturePath = f"{programPath}\\icon\\"
shortcutName = "CW Advanced.lnk"
startupPath = f"C:\\Users\\{getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"


# main window
class App(Tk):
    screenWidth = 1920
    screenHeigth = 1080

    screenHalfWidth = 960
    screenHalfHeigth = 540

    windows = []

    def __init__(self):
        super().__init__()

        # get screen resolution
        self.screenWidth = self.winfo_screenwidth()
        self.screenHeigth = self.winfo_screenheight()

        # get half width of screen dimensions in px
        self.screenHalfWidth = self.screenWidth / 2
        self.screenHalfHeigth = self.screenHeigth / 2

        # window is defined
        self.geometry("300x400")
        self.title("CW Advanced")
        self.iconbitmap(picturePath + "CenterWindow.ico")
        self.resizable(False, False)
        self.config(background="#242529")

        # elements in the main window are defined under here
        t = Label(self, text="by Marvin")
        t.config(font=("Arial", 7), background="#242529", foreground="#a6abb3")
        t.place(relx=0.0, rely=1.0, anchor="sw")

        t = Label(self, text="STRG + ALT + <")
        t.config(font=("Arial", 7), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.71, anchor="center")

        t = Label(self, text="All rights reserved")
        t.config(font=("Arial", 7), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=1, anchor="s")

        t = Label(self, text=f"Version {VERSION}")
        t.config(font=("Arial", 7), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.965, anchor="s")

        t = Label(self, text="CW Advanced")
        t.config(font=("Arial", 16), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.0, anchor="n")

        t = Label(
            self, text="All windows that are not\nminimized or maximized will be centered.")
        t.config(font=("Arial", 12), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.1, anchor="n")

        t = Button(self, text="Center", height=3,
                   width=15, command=self.centerWindows)
        t.config(font=("Arial", 12), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.6, anchor="center")

        t = Button(self, text="Center Self", width=9,
                   height=1, command=self.centerSelf)
        t.config(font=("Arial", 7), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=1.0, rely=1.0, anchor="se")

        # hotkey for fast centering
        add_hotkey("ctrl+alt+<", self.centerWindows)

        sleep(10)

        self.centerWindows()

    # function to center all windows except all maximized or minimized windows
    def centerWindows(self):
        # get all windows
        self.windows = getAllWindows()

        # check each window if not maximised or minimized and if name is not emty and if the window not self
        for i in self.windows:
            if not i.title == "":
                if not i.isMaximized:
                    if not i.title == "CW Advanced":
                        # get the half dimensions of the window in px
                        windowHalfWidht = i.width / 2
                        windowHalfHeight = i.height / 2

                        # calculates the center of the screen where the window should go
                        endWidth = self.screenHalfWidth - windowHalfWidht
                        endHeight = self.screenHalfHeigth - windowHalfHeight

                        # moves the window to the calculated position
                        i.moveTo(int(endWidth), int(endHeight))

    # centeres it self
    def centerSelf(self):
        # get all windows
        windowSelf = getWindowsWithTitle("CW Advanced")[0]

        # get the half dimensions of the window in px
        windowHalfWidht = windowSelf.width / 2
        windowHalfHeight = windowSelf.height / 2

        # calculates the center of the screen where the window should go
        endWidth = self.screenHalfWidth - windowHalfWidht
        endHeight = self.screenHalfHeigth - windowHalfHeight

        # moves the window to the calculated position
        windowSelf.moveTo(int(endWidth), int(endHeight))


# second window
class AboutUs(Tk):
    autostart = False

    def __init__(self):
        super().__init__()

        self.startTime = time()
        self.autostart = BooleanVar()

        # window is defined
        self.geometry("300x400")
        self.title("CW Advanced")
        self.iconbitmap(picturePath + "CenterWindow.ico")
        self.resizable(False, False)
        self.config(background="#242529")

        # elements in the second window are defined under here
        t = Label(self, text="Special thanks to BratWurst")
        t.config(font=("Arial", 16), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.03, anchor="n")

        t = Button(self, text="Discord", width=14,
                   height=2, command=self.discordLink)
        t.config(font=("Arial", 16), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.7, anchor="center")

        t = Button(self, text="Desktop Shortcut", width=14,
                   height=1, command=self.shortcutToDesktop)
        t.config(font=("Arial", 8), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=0.0, rely=1.0, anchor="sw")

        checkBox = Checkbutton(self, text="Autostart", variable=self.autostart,
                               background="#242529", foreground="#a6abb3")
        checkBox.place(relx=0.5, rely=1.0, anchor="s")

        self.createShortcut()

        # check starup folder if shortcut is in then select checkbox
        if isfile(startupPath + "\\" + shortcutName):
            checkBox.select()

        # after 10min enters function checkTime
        self.after(600000, self.onClose)
        # after 10min close window
        self.protocol("WM_DELETE_WINDOW", self.onClose)

    # Check wh|at the user want and where the program is
    def onClose(self):
        if isfile(startupPath + "\\" + shortcutName) and self.autostart.get():
            self.destroy()
        elif not isfile(startupPath + "\\" + shortcutName) and not self.autostart.get():
            self.destroy()
        elif isfile(startupPath + "\\" + shortcutName) and not self.autostart.get():
            try:
                remove(startupPath + "\\" + shortcutName)
                self.destroy()
            except:
                print("Error can not remove file")
                self.destroy()
        elif not isfile(startupPath + "\\" + shortcutName) and self.autostart.get():
            try:
                copy2(shortcutName, startupPath)
                self.destroy()
            except:
                print("Error can not copy and or paste file")
                self.destroy()
        else:
            self.destroy()

    # open link
    def discordLink(self):
        wbOpen("https://discord.gg/KbWUkUaSPv")

    # create a shortcut
    def createShortcut(self):
        if not isfile(shortcutName):
            try:
                shell = win32comDiscpatch("WScript.Shell")
                shortcut = shell.CreateShortCut(
                    f"{programPath}\\{shortcutName}")
                shortcut.Targetpath = f"{programPath}\\CW Advanced.exe"
                shortcut.IconLocation = f"{programPath}\\icon\\CenterWindow.ico"
                shortcut.WindowStyle = 1
                shortcut.save()
            except:
                print("Could not create shortcut")

    # copies the shortcut and paste it to the Desktop
    def shortcutToDesktop(self):
        if not isfile(shortcutName):
            self.createShortcut()

        copy2(shortcutName, join(join(environ["USERPROFILE"]), "Desktop"))


# start mainloop
if __name__ == "__main__":
    App().mainloop()
    AboutUs().mainloop()
