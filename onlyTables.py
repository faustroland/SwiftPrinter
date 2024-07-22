import win32api
import win32con
import pyautogui
import time
import ctypes
import pyperclip as pc #pip install pyperclip
import os
from zipfile import ZipFile 
from common import *

color_checking_coords: List[Tuple[int, int]] = [(10,10),(20,20)]

def waitForMenu(name,pos):
    for i in range(10):
        r = compareSquareAtPosition(name,pos[name])
        if r:
            break
        print("waiting for "+ name +" menu to appear")
        sleep(0.1)
    grabSquareAtPosition(name+"_fail",pos[name])

def click(position=[1920/2,1080/2],amount=1):
    getActiveWindow()
    x,y=position
    x=int(x)
    y=int(y)
    win32api.SetCursorPos((x,y))
    time.sleep(0.025)
    for loop in range(0,amount):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        time.sleep(0.025)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
        time.sleep(0.025)

def rightclick(position=[1920/2,1080/2],amount=1):
    getActiveWindow()
    x,y=position
    x=int(x)
    y=int(y)
    win32api.SetCursorPos((x,y))
    for loop in range(0,amount):
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)
        time.sleep(0.025)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)
        time.sleep(0.01)


def kp(key,length=0.1,wait=0.1):
    getActiveWindow()
    win32api.keybd_event(key,0,0,0)  
    time.sleep(length)
    win32api.keybd_event (key,0, win32con.KEYEVENTF_KEYUP, 0) # key is released    
    if wait:
        time.sleep(wait)

def escape():
    getActiveWindow()
    kp(0x1B,0.1,0.1)

def enter():
    getActiveWindow()
    kp(0x0D,0.1,0.1)

def makerPen():
    getActiveWindow()
    win32api.keybd_event(0x31,0,0,0)
    sleep(0.5)
    win32api.SetCursorPos((int(1920/2)-200,int(1080/2)+100))
    sleep(0.1)
    win32api.keybd_event (0x31,0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(0.1)

def load_positions(file_path):
    settings = {}
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line by '='
            key, value = line.strip().split('=')
            # Store key-value pair in dictionary
            settings[key.strip()] = tuple(map(int, value.strip().split(',')))
    return settings

def load_settings(file_path):
    settings = {}
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line by '='
            key, value = line.strip().split('=')
            settings[key.strip()] = value
    return settings

def load_colors(file_path):
    colors = []
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line by '='
            colors.append(line.strip())
    return colors

def saveRoomFromMP(f):
    getActiveWindow()
    kp(0x46) #F
    sleep(4)
    click(pos["MP_ThisRoom"])
    sleep(4)
    click(pos["SaveRoom"])
    sleep(4)
    click(pos["SaveDescription"])
    paste(str(f))
    sleep(4)
    click(pos["Save"])

def split_file(input_file, chunk_size):
    with open(input_file, 'r',encoding="utf-8") as file:
        lines = file.readlines()
    l = 0
    start=0
    end = 0
    chunks = []
    for i in range(len(lines)):
        lines[i]=lines[i].strip()
        l = l + len(lines[i])
        if l>chunk_size:
            chunks.append((lines[start:i]))
            start = i
            l=0
    chunks.append((lines[start:-1]))
    textchunks = []
    lenchunks = len(chunks)
    rnd = "0"
    for i in range(len(chunks)):
        if i==0:
            chunks[0][0] = chunks[0][0] + ";"+str(lenchunks)
            q = chunks[0][0].split(";")
            rnd = q[4]
            print(chunks[0][0])
            print(rnd)
            textchunks.append("\n".join(chunks[0]))
        else:
            chunks[i].insert(0,str(i)+";"+rnd)
            textchunks.append("\n".join(chunks[i]))
            
    
    return textchunks

def getStatus(pos):
    getActiveWindow
    x,y = pos["STATUS"]
    image = ImageGrab.grab()
    RGB = image.load()[x,y]
    R,G,B=RGB
    R=int(round((R-13)/2**5))
    G=int(round((G-13)/2**5))
    B=int(round((B-13)/2**5))
    val=R<<6|G<<3|B
    return(val)

def getColorAt(pos):
    getActiveWindow
    x,y = pos
    image = ImageGrab.grab()
    RGB = image.load()[x,y]
    R,G,B=RGB
    val=R<<16|G<<8|B
    return(val)

def press(key):
    kp(VK_CODE[key])

def dropMakerPen(delay):
    press("z")
    sleep(delay)

def makerPenMenu(delay):
    press("f")
    sleep(delay)


def enterData():
    pos=load_positions("positions.txt")
    
    settings = load_settings("settings.txt")
    c_delay = float(settings["color_import_delay"])
    b_delay = float(settings["button_delay"])
    data_chunk_size = int(settings["data_chunk_size"])
    
    path=os.getcwd()
    files=os.listdir(path)
    files=[f for f in files if ".png.zip" in f]
    print(files)
    #prepared for batch importing, currently cut out
    for f in files[0:1]:
        print(str(f))
        print("extracting: ",f)
        with ZipFile(f,"r") as zObject:
            zObject.extractall(path)

        chunks = split_file("image_data.txt",data_chunk_size)
        colors = load_colors("image_hex.txt")
        lenchunks=len(chunks)
        lencolors=len(colors)
        print(f"Importing {lencolors} colors and {lenchunks} TABLES")
        getActiveWindow()


        dropMakerPen(1)
        makerPen()
        makerPenMenu(1)
#        waitForMenu("MP_tools",pos)
        click(pos["MP_tools"])
        sleep(1)
        click(pos["RecolorButt"])
        escape()


        iC = 0
        circuitsClicked = False
        
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
            sleep(2)
            click(pos["DATA_field"])                    
            sleep(4)
            ctrlA()
            sleep(1)
            ctrlA()
            sleep(1)
            kp(0x2E)
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


def waitForDone():    
    while(True):
        if (getStatus(pos) == 200):
            break
        else:
            sleep(1)


        #wait for the save
        while(True):
            if not getStatus(pos)==300:
                sleep(1)
            else:
                sleep(6)
                break
        #wait for the next room load
        while(True):
            if not getStatus(pos)==300:
                sleep(1)
            else:
                sleep(10)
                break



#Script States
ENTERING_DATA = 0
PRINTING = 310
SAVING = 2
UNDECIDED = 99
ERROR = 666
DONE = 1000

#Printer States
ROOM_LOADED = 300
SEATED = 0
PRINTING = 310
PRINTING_DONE = 200


if __name__ == "__main__":
    pos=load_positions("positions.txt")
   
    status = getStatus(pos)
    state = UNDECIDED


    while(True):
        getActiveWindow()
        if state == UNDECIDED:
            print("UNDECIDED")
            sleep(1)
            status = getStatus(pos)
            if status==SEATED:
                state = ENTERING_DATA
                continue
            elif status == ROOM_LOADED:
                state = UNDECIDED
                continue
            elif status ==PRINTING:
                state = PRINTING
                continue
            elif status == PRINTING_DONE:
                state = SAVING
                continue
            sleep(10)
            
        if state == ENTERING_DATA:
            print("ENTERING DATA")
            enterData()
            while(True):
                if getStatus(pos)== PRINTING:
                    state = PRINTING
                    break
                sleep(1)


        if state == PRINTING:
            print("PRINTING")
            status = getStatus(pos)
            print(state)
            if status == PRINTING_DONE:
                 state = SAVING
                 continue

            sleep(10)
            state = UNDECIDED

        if state == SAVING:
            print("SAVING")
            saveRoomFromMP("DONE")
            while(True):
                if getStatus(pos)== ROOM_LOADED:
                    state = DONE
                    break
                sleep(1)

        if state == ERROR:
            print("ERROR")
            ## TODO solve error
            break

        if state == DONE:
            print("DONE")
            break

        sleep(1)
            
    input("Press ENTER to exit")
