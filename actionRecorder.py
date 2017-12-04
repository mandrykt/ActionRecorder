import tkinter as tk
import pyautogui
import win32gui
import time
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pywinauto.findwindows import find_window
from pywinauto.win32functions import SetForegroundWindow
import subprocess
from pywinauto import application
import keyboard
import win32api

state_left = win32api.GetKeyState(0x01)
state_right = win32api.GetKeyState(0x02)

mouseClickLocations = []

while True:
    a = win32api.GetKeyState(0x01)
    b = win32api.GetKeyState(0x02)

    if a != state_left:
        state_left = a
        #print(a)
        if a < 0:
            #print("left button clicked")
            mouseClickLocations.append(pyautogui.position())
            #print(mouseClickLocations)
        #else:
        #    print("left button released")

    #if keyboard.is_pressed('a') and mouseClickLocations != []:
    #    for locations in mouseClickLocations:
            
            ###move mouse to location and open file
            #win32api.SetCursorPos((locations[0], locations[1]))

            ###open the file
            #pyautogui.click(clicks=2)
     #       print(locations)

    
    if keyboard.is_pressed('a'):    ###try to get an actual key binding
        if mouseClickLocations == []:
            location = pyautogui.position()
            
            ###open the file
            pyautogui.click(clicks=2)

            ###wait for file to open
            time.sleep(1)

            ###get window title (ie. file name)
            windowTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            print(windowTitle)

            filename = windowTitle.split(".pdf")[0] + ".pdf"
            path = "C:/Users/Taras/Desktop/PDFs/"
            fp = open(path + filename, 'rb')
            parser = PDFParser(fp)
            doc = PDFDocument(parser)

            ###PDF title
            pdfTitle = (doc.info[0]['Title']).decode("utf-8")
            print(pdfTitle)  # The "Info" metadata

            ###close the PDF window
            subprocess.call(["taskkill", "/f", "/im", "ACROBAT.EXE"])

            #restore Notepad and bring it to the foreground
            app = application.Application()
            app.connect(title_re=".*%s.*" % 'Notepad')
            app_dialog = app.top_window_()
            app_dialog.Minimize()
            app_dialog.Restore()

            #app.Notepad.Edit.set_edit_text(u"I am typing s\xe4me text to Notepad\r\n\r\n")
            app.Notepad.Edit.type_keys(u"{END}{ENTER}" + pdfTitle, with_spaces = True)

            app_dialog.Minimize()

        else:
            for locations in mouseClickLocations:
                win32api.SetCursorPos((locations[0], locations[1]))
                pyautogui.click(clicks=2)
                
                ###wait for file to open
                time.sleep(1)

                ###get window title (ie. file name)
                windowTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                print(windowTitle)

                filename = windowTitle.split(".pdf")[0] + ".pdf"
                path = "C:/Users/Taras/Desktop/PDFs/"
                fp = open(path + filename, 'rb')
                parser = PDFParser(fp)
                doc = PDFDocument(parser)

                ###PDF title
                pdfTitle = (doc.info[0]['Title']).decode("utf-8")
                print(pdfTitle)  # The "Info" metadata

                ###close the PDF window
                subprocess.call(["taskkill", "/f", "/im", "ACROBAT.EXE"])

                #restore Notepad and bring it to the foreground
                app = application.Application()
                app.connect(title_re=".*%s.*" % 'Notepad')
                app_dialog = app.top_window_()
                app_dialog.Minimize()
                app_dialog.Restore()

                #app.Notepad.Edit.set_edit_text(u"I am typing s\xe4me text to Notepad\r\n\r\n")
                app.Notepad.Edit.type_keys(u"{END}{ENTER}" + pdfTitle, with_spaces = True)

                app_dialog.Minimize()
