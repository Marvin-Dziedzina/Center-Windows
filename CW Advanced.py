from pygetwindow import getAllWindows
from tkinter import *
from keyboard import add_hotkey
import webbrowser
from shutil import copy2
import win32com.client
import os
import time

programPath = os.path.dirname(os.path.realpath(__file__))
picturePath = programPath + "\\icon\\"
shortcutName = "CW startup.lnk"
startupPath = "C:\\Users\\" + os.getlogin() + \
    "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"


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
        self.windows = getAllWindows()

        for i in self.windows:
            if i.title == "CW Advanced":
                # get the half dimensions of the window in px
                windowHalfWidht = i.width / 2
                windowHalfHeight = i.height / 2

                # calculates the center of the screen where the window should go
                endWidth = self.screenHalfWidth - windowHalfWidht
                endHeight = self.screenHalfHeigth - windowHalfHeight

                # moves the window to the calculated position
                i.moveTo(int(endWidth), int(endHeight))


# second window
class AboutUs(Tk):
    autostart = ""
    startTime = 0
    actualTime = 0

    def __init__(self):
        super().__init__()

        self.startTime = time.time()
        self.autostart = BooleanVar()

        # window is defined
        self.geometry("300x400")
        self.title("CW Advanced")
        self.iconbitmap(picturePath + "CenterWindow.ico")
        self.resizable(False, False)
        self.config(background="#242529")

        # elements in the second window are defined under here
        t = Label(self, text="Special thanks to Moritz\naka Brat__Wurst")
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
        if os.path.isfile(startupPath + "\\" + shortcutName):
            checkBox.select()

        # after 10sec enters function checkTime
        self.after(10000, self.checkTime)

        self.protocol("WM_DELETE_WINDOW", self.onClose)

    # Check waht the user want and where the program is
    def onClose(self):
        if os.path.isfile(startupPath + "\\" + shortcutName) and self.autostart.get():
            self.destroy()
        elif not os.path.isfile(startupPath + "\\" + shortcutName) and not self.autostart.get():
            self.destroy()
        elif os.path.isfile(startupPath + "\\" + shortcutName) and not self.autostart.get():
            try:
                os.remove(startupPath + "\\" + shortcutName)
                self.destroy()
            except:
                print("Error can not remove file")
                self.destroy()
        elif not os.path.isfile(startupPath + "\\" + shortcutName) and self.autostart.get():
            try:
                copy2(shortcutName, startupPath)
                self.destroy()
            except:
                print("Error can not copy and or paste file")
                self.destroy()
        else:
            self.destroy()

    # close window after 10min
    def checkTime(self):
        self.actualTime = time.time() - self.startTime
        if self.actualTime >= 600:
            self.destroy()
        self.after(10000, self.checkTime)

    # open link
    def discordLink(self):
        webbrowser.open("https://discord.com/invite/pNpyYAhz5U")

    # create a shortcut
    def createShortcut(self):
        if not os.path.isfile(shortcutName):
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(
                    programPath + "\\" + shortcutName)
                shortcut.Targetpath = programPath + "\\" + "CW Advanced.exe"
                shortcut.IconLocation = programPath + "\\icon\\CenterWindow.ico"
                shortcut.WindowStyle = 1
                shortcut.save()
            except:
                print("Could not create shortcut")

    # copies the shortcut and paste it to the Desktop
    def shortcutToDesktop(self):
        if not os.path.isfile(shortcutName):
            self.createShortcut()
        copy2(shortcutName, os.path.join(
            os.path.join(os.environ["USERPROFILE"]), "Desktop"))


# start mainloop
if __name__ == "__main__":
    App().mainloop()
    AboutUs().mainloop()
