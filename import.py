import win32api
import win32con
import pyautogui
import time
import ctypes
import pyperclip as pc #pip install pyperclip
import os
from zipfile import ZipFile 
from common import *
from colors import *
from tables import *

color_checking_coords: List[Tuple[int, int]] = [(10,10),(20,20)]

def waitForMenu(name,pos):
    for i in range(10):
        r,exceeded_points = compareSquareAtPosition(name,pos[name])
        if r:
            return True
        print("waiting for "+ name +" menu to appear")
        sleep(0.1)
    grabSquareAtPosition(name+"_fail",pos[name])
    return False

def scrollup(pos):
    getActiveWindow()
    x,y=pos["Recolor_Tool_Setings"]
    x=int(x)
    y=int(y)
    win32api.SetCursorPos((x,y))
    time.sleep(0.025)
    for i in range(10):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, x, y, 1, 0)
        time.sleep(0.025)


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

def load_positions(file_path):
    settings = {}
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line by '='
            key, value = line.strip().split('=')
            # Store key-value pair in dictionary
            settings[key.strip()] = tuple(map(int, value.strip().split(',')))
    return settings

def enterData(startat):

#    res_affair_17 = [1920,1366,1600,2560,3840]
#    res_affair_16 = [1440,1280]
    pos=load_positions("positions.txt")   
    settings = load_settings("settings.txt")
    user32 = ctypes.windll.user32
    my_res = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    pos_res = list(map(int,settings["position_resolution"].split(",")))
    
    if my_res[0]!=pos_res[0]:
        print("Recalculating coordinates")
        for i in pos:
            x,y = pos[i]
            new_x = int(x*my_res[0]/pos_res[0])
            new_y = int(y*my_res[1]/pos_res[1])
            pos[i] = [new_x,new_y]
 
        

    
    c_delay = float(settings["color_import_delay"])
    b_delay = float(settings["button_delay"])
    data_chunk_size = int(settings["data_chunk_size"])
    path=os.getcwd()
    files=os.listdir(path)
    files=[f for f in files if ".png.zip" in f]
    name = ""
    print(files)
    if 0==len(files):
        print("No file for import")
        return False
    f = files[0]
    print(str(f))
    print("extracting: ",f)
    name=f
    with ZipFile(f,"r") as zObject:
        zObject.extractall(path)

    chunks = split_file("image_data.txt",data_chunk_size)
    colors = load_colors("image_hex.txt")
    lenchunks=len(chunks)
    lencolors=len(colors)
    print(f"Importing {lencolors} colors and {lenchunks} TABLES")
#    getActiveWindow()
    
    importColors(pos,settings,c_delay,b_delay,colors,startat)
    saveRoomFromMP("Colors",pos)
    while(True):
        if (getStatus(pos) == 300):
            sleep(8)
            break
        else:
            escape()
            sleep(1)
    importTables(pos,settings,c_delay,b_delay,data_chunk_size,chunks)
    return "tampName"


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
    yes = {'yes','y', 'ye', ''}
    no = {'no','n'}
    choice = input("Is this a shirt? [Y/n]: ").lower()
    startat=(input("Start at marker [0]: ").lower())
    if len(startat)==0:
        startat=0
    else:
        startat=int(startat)
    if choice in yes:
       shirt=True
    elif choice in no:
       shirt=False
    else:
       sys.stdout.write("Please respond with 'yes' or 'no'")
       
    pos=load_positions("positions.txt")
    settings = load_settings("settings.txt")
    my_res = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    pos_res = list(map(int,settings["position_resolution"].split(",")))
    
    if my_res[0]!=pos_res[0]:
        print("Recalculating coordinates")
        
        for i in pos:
            x,y = pos[i]
            new_x = int(x*my_res[0]/pos_res[0])
            new_y = int(y*my_res[1]/pos_res[1])
            pos[i] = [new_x,new_y]
            
    #shirt=.lower.contains("Y")
   
 #   status = getStatus(pos)
 #   state = UNDECIDED

    name = enterData(startat)
    

    while(True):
        status = getStatus(pos)
       # print(state)
        if status == PRINTING_DONE:
            break

    if not shirt:
        saveRoomFromMP(name,pos)

            
    input("Press ENTER to exit")
