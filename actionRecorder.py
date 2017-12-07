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
import mouse
from win32gui import GetWindowText, GetForegroundWindow
from ctypes import windll
from msvcrt import getch
from pywinauto.win32functions import SetForegroundWindow, ShowWindow
import sys

mouseClickLocations = []
mouseEvents = []
keyboardEvents = []
mouseStatus = False
keyboardStatus = False
mUnhooked = False
kUnhooked = False

def filterEvents(eventList):

    eventDescriptions = []

    ### find first pdf file location in action list
    firstPDFIndex = 0
    for item in eventList:
        if ".pdf" in item[3]:
            firstPDFIndex = eventList.index(item)
            break

    ### remove occurances of 'Program Manager' that occur before the first pdf file
    pmIndices = []
    for i in range(len(eventList)):
        if i < firstPDFIndex and eventList[i][3] == "Program Manager":
            pmIndices.append(i)

    eventList = [i for j, i in enumerate(eventList) if j not in pmIndices]

    ### remove 'middle'
    eventList = list(filter(lambda item: item[1] != "middle", eventList))
    
    ### open PDF file. Do we need to care about double click on this action? Will need to use coordinates on this action
    eventList = list(filter(lambda item: item[3] != "PDFs", eventList))
    eventDescriptions.append("OPEN PDF")

    FIRST_TITLE = ''
    FIRST_TITLE_INDEX = 0
    for item in eventList:
        if ".pdf" in item[3]:
            FIRST_TITLE = item[3]
            FIRST_TITLE_INDEX = eventList.index(item)
            break

    ### might need to change this bit...doesnt work if the first item in eventList isnt the pdf title element
    LAST_OCCURANCE = 0    
    for item in eventList[FIRST_TITLE_INDEX:]:
        if item[3] == FIRST_TITLE: # or LAST_OCCURANCE == 0:
            LAST_OCCURANCE = eventList.index(item)
        else:
            break
    
    titleSet = [item[3] for item in eventList[FIRST_TITLE_INDEX:LAST_OCCURANCE + 1]]
    
    if len(set(titleSet)) == 1:
        eventDescriptions.append("COPY TITLE")
        eventList = eventList[LAST_OCCURANCE + 1:]

    ### are we pasting a title?
    size = len(eventList)
    firstEventList = list(filter(lambda item: "Notepad" not in item[3], eventList))
    otherEventList = list(filter(lambda item: "Program Manager" not in item[3], eventList))
    if size > len(firstEventList): # or size > len(otherEventList):
        eventList = firstEventList
        eventDescriptions.append("PASTE TITLE")
    elif size > len(otherEventList):
        eventList = otherEventList
        eventDescriptions.append("PASTE TITLE")

    ### are we closing the PDF?
    for item in eventList:
        if item[3] == FIRST_TITLE or item[3] == "Adobe Acrobat Pro DC" or ".pdf" in item[3]:
            eventDescriptions.append("CLOSE PDF")
            break

    return eventDescriptions

def closePDF():
    ### close the PDF window
    subprocess.call(["taskkill", "/f", "/im", "ACROBAT.EXE"])

def openPDF():
    ### open the file
    pyautogui.click(clicks=2)

    ### wait for file to open
    time.sleep(1)

    ### get window title (ie. filename)
    windowTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())

    ### open the file
    filename = windowTitle.split(".pdf")[0] + ".pdf"
    path = "C:/Users/Taras/Desktop/PDFs/"
    fp = open(path + filename, 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument(parser)

    return doc

def getTitle(document):
    return (document.info[0]['Title']).decode("utf-8")

def pasteTitle(pdfTitle):
    if pdfTitle != None:
        ### restore notepad
        app = application.Application()
        app.connect(title_re=".*%s.*" % 'Notepad')
        app_dialog = app.top_window_()
        app_dialog.Minimize()
        app_dialog.Restore()

        ### add the text
        app.Notepad.Edit.type_keys(u"{END}{ENTER}" + pdfTitle, with_spaces = True)
        app_dialog.Minimize()
    else:
        print("error: couldn't paste title")
        sys.exit()


def endOrRestart():
    userInput = input("Press 'q' to quit or 'r' to restart\n")
    if userInput == 'q':
        exit()
    elif userInput == 'r':
        start()
    else:
        userInput = input("Please press 'q' or 'r'\n")
        
def getMouseClicks(event):
    global mUnhooked
    global mouseClickLocations
    if keyboard.is_pressed("z"):
        if mouse.is_pressed(button="left"):
            mouseClickLocations.append(mouse.get_position())
            mouseClickLocations = list(set(mouseClickLocations))
    if mouse.is_pressed(button="middle"):
        print("unhooking mouse")
        mouse.unhook(getMouseClicks)
        mUnhooked = True

        if kUnhooked and mUnhooked:
            endOrRestart()

    
def userActions(event):
    allEvents = mouseEvents + keyboardEvents
    sortedEvents = sorted(allEvents, key=lambda x: x[-1])
    finalEvents = filterEvents(mouseEvents)

    global mouseClickLocations
    
    if keyboard.is_pressed('a'):
        document = None
        title = ''
        for item in finalEvents:
            if item == "OPEN PDF":
                document = openPDF()
            elif item == "COPY TITLE":
                title = getTitle(document)
            elif item == "PASTE TITLE":
                pasteTitle(title)
            elif item == "CLOSE PDF":
                closePDF()
    if keyboard.is_pressed('esc'):
        print("unhooking keyboard")
        keyboard.unhook(userActions)
        kUnhooked = True

        if kUnhooked and mUnhooked:
            endOrRestart()
   

    if keyboard.is_pressed('space'):
        for location in mouseClickLocations:
            win32api.SetCursorPos((location[0], location[1]))
            document = None
            title = ''
            for item in finalEvents:
                if item == "OPEN PDF":
                    document = openPDF()
                elif item == "COPY TITLE":
                    title = getTitle(document)
                elif item == "PASTE TITLE":
                    pasteTitle(title)
                elif item == "CLOSE PDF":
                    closePDF()
        
        ### clear the mouseClickLocations
        mouseClickLocations = []


def MainLoop():
    print("Press 'a' to repeat once or 'space' to repeat multiple times")
    keyboard.hook(userActions)
    mouse.hook(getMouseClicks)
 
def OnMouseEvent(event):
    ### unhook the mouse
    if mouse.is_pressed(button='middle'):
        print("unhooking mouse")
        mouse.unhook(OnMouseEvent)
        global mouseStatus
        mouseStatus = True

        if (mouseStatus and keyboardStatus):
            MainLoop()
                    
    if type(event).__name__ == 'ButtonEvent' and event.event_type == 'down':   ### is this a button click?
        currentTime = event.time
        previousTime = 0
        
        ### grab the last event time if available
        if mouseEvents != []:
            previousTime = mouseEvents[-1][4]
        
        button = event.button
        location = mouse.get_position()
        eventTime = event.time
        
        ### need to also get the currently focused window
        window = GetWindowText(GetForegroundWindow())
        
        ### this is just to get the name of the window if we click on an out of focus window
        if window == '' and not 0 <= location[0] <= 45 and not 1077 <= location[1] <= 1043:      ### start button on windows
            ### simulate a button click
            pyautogui.click(clicks=1)
        

        if currentTime - previousTime < 0.40:
            ### remove the previous event
            if (currentTime - previousTime < 0.006):        ### this must be a simulated click
                mouseEvents[-1] = ('mouse', button, location, window, eventTime)
            else:
                del mouseEvents[-1]
                mouseEvents.append(('mouse', 'double', location, window, eventTime))
        else:
            mouseEvents.append(('mouse', button, location, window, eventTime))
    
    return True

def OnKeyboardEvent(event):
    if keyboard.is_pressed('esc'):
        print("unhooking keyboard")
        keyboard.unhook(OnKeyboardEvent)
        global keyboardStatus
        keyboardStatus = True

        if (mouseStatus and keyboardStatus):
            MainLoop()
   
    elif type(event).__name__ == 'KeyboardEvent' and event.event_type == 'down':
        key = event.name
        eventType = 'down'
        window = GetWindowText(GetForegroundWindow())
        eventTime = event.time
        keyboardEvents.append(('key', key, eventType, window, eventTime))
        
    return True

def start():
    print("Press 'q' to begin")
    while True:
        if keyboard.is_pressed('q'):
            mouse.hook(OnMouseEvent)
            keyboard.hook(OnKeyboardEvent)
            break

if __name__ == "__main__":
    start()
