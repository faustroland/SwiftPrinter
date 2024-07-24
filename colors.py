import win32api
import win32con
import pyautogui
import time
import ctypes
import pyperclip as pc #pip install pyperclip
import os
from zipfile import ZipFile 
from common import *

def importColors(pos,settings,c_delay,b_delay,colors):
    getActiveWindow()


    dropMakerPen(1)
    makerPen() #open
    makerPenMenu(1)
    click(pos["MP_tools"]) # click the Maker Pen "Tools" button
    sleep(1)
    click(pos["MP_configure"]) # yes, a butt joke
    scrollup(pos)
    escape() #close


    iC = 0
    circuitsClicked = False
    # Recoloring 
    for color in colors:
        print("")  
        print(color)
        getActiveWindow()
        print(iC)
        if iC < 0: # for debuggung to skip some colors
            iC=iC+1
            continue

        while(True):  # here is handled the transition to the next marker
            position = getStatus(pos)
            print("My position:",position)
            if position < iC: 
                rightclick([1920/2,1080/2])
                sleep(0.2)
            elif position > iC: ## this part handles an "overclick" when you are ahead, it resets the room and starts over 
                makerPenMenu(1) #open
                click(pos["MP_tools"])
                sleep(1)
                click(pos["MP_configure"])
                sleep(1)
                if not circuitsClicked:
                    click(pos["Circuits"])
                    sleep(1)
                    circuitsClicked=True
                click(pos["RoomReset"])
                sleep(0.1)
                kp(0x46) #Press F, close maker pen
                sleep(0.1)
                kp(0x5A) #Z drop maker penn
                sleep(0.1)
                kp(0x20) # space, jump out of the seat
                sleep(0.1)
                sleep(6)
                makerPen()
                sleep(1)
                kp(0x46) #F
                sleep(1)
                click(pos["MP_tools"])
                sleep(1)
                click(pos["RecolorButt"])
                escape()
                for qq in range(iC):
                    rightclick([1920/2,1080/2])
                    sleep(0.2)
            else:
                break
            
        click([1920/2,1080/2],1)
        
        waitForMenu("MARKER_MULTICOLOR_CONFIG_MENU",pos)
        waitForMenu("MULTICOLOR_COLOR_BUTTON",pos)
        click(pos["MULTICOLOR_COLOR_BUTTON"])

        waitForMenu("Custom",pos)
        click(pos["Custom"])
        while not waitForMenu("CUSTOM_COLOR_MENU_HEADER",pos):
            click(pos["Custom"])
            sleep(c_delay)
        click(pos["Custom_Input"],3)
        sleep(c_delay)
        ctrlA()
        ctrlA()
        sleep(0.01)
        paste(color)
        sleep(0.1*c_delay)
        enter()
        makerPenMenu(0.1*c_delay) #close 

        iC=iC+1

    sleep(c_delay)
    rightclick([1920/2,1080/2],1)
