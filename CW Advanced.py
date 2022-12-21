from pygetwindow import getAllWindows, getWindowsWithTitle
from tkinter import Label, Button, Tk, BooleanVar, PhotoImage, Checkbutton
from keyboard import add_hotkey
from webbrowser import open as wbOpen
from shutil import copy2
from win32com.client import Dispatch as win32comDiscpatch
from os import getlogin, environ, remove
from os.path import dirname, realpath, isfile, join, getsize
from time import time

# Version with in this format
# X.Y.Z
# X = Major Version
# Y = Minor Version
# Z = Patch
#
VERSION = "1.2.0"

programPath = dirname(realpath(__file__))
picturePath = f"{programPath}\\icon\\"
shortcutName = "CW Advanced.lnk"
startupPath = f"C:\\Users\\{getlogin()}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
logFilePath = "log.txt"


# main window
class App(Tk):
    screenWidth = 1920
    screenHeigth = 1080

    screenHalfWidth = 960
    screenHalfHeigth = 540

    windows = []

    def __init__(self):
        super().__init__()

        self.log("__init__", "In init func")

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
        t = Label(self, text="by Mixaffected")
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

        self.photo = PhotoImage(file=f"{picturePath}settings.png")
        t = Button(self, text="Settings",
                   command=self.openSettingsWindow, image=self.photo)
        t.config(font=("Arial", 12), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=0.98, rely=0.02, anchor="ne")

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

        self.createShortcut()

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

                        # get screen resolution
                        self.screenWidth = self.winfo_screenwidth()
                        self.screenHeigth = self.winfo_screenheight()

                        # get half width of screen dimensions in px
                        self.screenHalfWidth = self.screenWidth / 2
                        self.screenHalfHeigth = self.screenHeigth / 2

                        # calculates the center of the screen where the window should go
                        endWidth = self.screenHalfWidth - windowHalfWidht
                        endHeight = self.screenHalfHeigth - windowHalfHeight

                        self.log("screenWidth", self.screenWidth)
                        self.log("screenHeigth", self.screenHeigth)
                        self.log("endWidth", endWidth)
                        self.log("endHeight", endHeight)

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
        endWidth = self.screenHalfWidth / 3 - windowHalfWidht
        endHeight = self.screenHalfHeigth / 1.8 - windowHalfHeight

        # moves the window to the calculated position
        windowSelf.moveTo(int(endWidth), int(endHeight))

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
                self.log("CreatedShortcut", "Shortcut created")
            except:
                print("Could not create shortcut")

    def openSettingsWindow(self):
        self.log("Open settings", "Open Settings")
        Settings().mainloop()

    def log(self, title="", message=""):
        if isfile("log.txt"):
            fileSize = getsize(logFilePath)
            print(fileSize)
            message = f"{title}: {message} |{fileSize}"

            if fileSize / 1024 >= 5000:
                contents = []

                with open(logFilePath, "r") as log:
                    contents = log.readlines()

                count = 0
                for idx, e in enumerate(contents):
                    if count <= 30:
                        count += 1
                        continue

                    contents.pop(idx)

                with open(logFilePath, "w") as log:
                    log.writelines(contents)

        try:
            with open(logFilePath, "x") as log:
                log.write(f"{message}\n")
        except:
            with open(logFilePath, "a") as log:
                log.write(f"{message}\n")


# options window
class Settings(Tk):
    autostart = False

    def __init__(self):
        super().__init__()

        self.log("Setting init", "In settings init")

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

        # check starup folder if shortcut is in then select checkbox
        if isfile(f"{startupPath}\\{shortcutName}"):
            checkBox.select()

        # after 10min enters function checkTime
        self.after(120000, self.onClose)
        # after 10min close window
        self.protocol("WM_DELETE_WINDOW", self.onClose)

    # Check wh|at the user want and where the program is
    def onClose(self):
        if isfile("f{startupPath}\\{shortcutName}") and self.autostart.get():
            self.destroy()
        elif not isfile(f"{startupPath}\\{shortcutName}") and not self.autostart.get():
            self.destroy()
        elif isfile(f"{startupPath}\\{shortcutName}") and not self.autostart.get():
            try:
                remove(f"{startupPath}\\{shortcutName}")
                self.destroy()
            except:
                print("Error can not remove file")
                self.destroy()
        elif not isfile(f"{startupPath}\\{shortcutName}") and self.autostart.get():
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

    # copies the shortcut and paste it to the Desktop
    def shortcutToDesktop(self):
        if not isfile(shortcutName):
            self.createShortcut()

        copy2(shortcutName, join(join(environ["USERPROFILE"]), "Desktop"))

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

    def log(self, title="", message=""):
        if isfile("log.txt"):
            fileSize = getsize(logFilePath)
            print(fileSize)
            message = f"{title}: {message} |{fileSize}"

            if fileSize / 1024 >= 5000:
                contents = []

                with open(logFilePath, "r") as log:
                    contents = log.readlines()

                count = 0
                for idx, e in enumerate(contents):
                    if count <= 30:
                        count += 1
                        continue

                    contents.pop(idx)

                with open(logFilePath, "w") as log:
                    log.writelines(contents)

        try:
            with open(logFilePath, "x") as log:
                log.write(f"{message}\n")
        except:
            with open(logFilePath, "a") as log:
                log.write(f"{message}\n")


# start mainloop
if __name__ == "__main__":
    App().mainloop()
