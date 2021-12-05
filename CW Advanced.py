import pygetwindow as gw
from tkinter import *
import keyboard
import webbrowser

# change for build
# picturePath = ".\\CenterWindow\\icon\\" # the relative path on my computer
picturePath = ".\\icon\\"  # the relative path if in the build is an folder


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
        keyboard.add_hotkey("ctrl+alt+<", self.centerWindows)

    # function to center all windows except all maximized or minimized windows
    def centerWindows(self):
        # get all windows
        self.windows = gw.getAllWindows()

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
        self.windows = gw.getAllWindows()

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
    def __init__(self):
        super().__init__()
        # window is defined
        self.geometry("300x400")
        self.title("CW Advanced")
        self.iconbitmap(picturePath + "CenterWindow.ico")
        self.resizable(False, False)
        self.config(background="#242529")

        # elements in the second window are defined under here
        t = Label(self, text="Special thanks to Moritz")
        t.config(font=("Arial", 16), background="#242529", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.03, anchor="n")

        t = Button(self, text="Discord", width=14,
                   height=2, command=self.discordLink)
        t.config(font=("Arial", 16), background="#373a40",
                 activebackground="#292c30", foreground="#a6abb3")
        t.place(relx=0.5, rely=0.7, anchor="center")

    # open link
    def discordLink(self):
        webbrowser.open("https://discord.com/invite/pNpyYAhz5U")


# start mainloop
if __name__ == "__main__":
    App().mainloop()
    AboutUs().mainloop()
