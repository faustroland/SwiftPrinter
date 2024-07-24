import win32api
import win32con
import pyautogui
import time
import ctypes
import pyperclip as pc #pip install pyperclip
import os
from zipfile import ZipFile 
from common import *

def importTables(pos,settings,c_delay,b_delay,data_chunk_size,chunks):    

        kp(0x5A) #Z
        sleep(2)
        for i in range(len(chunks)):
            print(i)
            getActiveWindow()
            
            for y in range(i):
                rightclick([1920/2,1080/2],1)
                sleep(0.5)
                
            makerPen()
            sleep(1)
            kp(0x46) #F
            sleep(2)
            click(pos["MP_tools"])
            sleep(2)
            click(pos["MP_configure"])
            sleep(1)
            escape()
            sleep(1)
            click()
            sleep(2)
            click(pos["Edit_Data_Table_Button"])
            sleep(3)
            click(pos["DataTable_Enter_BOX"])
            while not waitForMenu("DATA_TABLE_LOADED",pos):
                sleep(1)
            click(pos["DATA_field"])                    
            sleep(4)
            ctrlA()
            sleep(1)
            ctrlA()
            sleep(1)
            kp(0x2E) #delete
            sleep(1)
            paste(chunks[i])
            print("Sleep 45")
            sleep(45)
            click(pos["GENERATE"])
            sleep(10)

            while(True):
                if (getStatus(pos) == 300):
                    sleep(8)
                    break
                else:
                    sleep(0.1)
        kp(0x5A) #Z
        sleep(6)    
        makerPen()
        sleep(1)
        kp(0x46) #F
        click(pos["MP_tools"])
        sleep(2)
        click(pos["MP_configure"])
        sleep(1)
        click(pos["Circuits"])
        sleep(1)
        click(pos["TestEvent"])
        kp(0x46) #F
        # move data file to done folder
        src = path + "\\" + f
        dest = path + "\\done\\" + f
        os.rename(src, dest)
