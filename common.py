import win32api, win32con
import time
import datetime
import logging
import sys
import time
from typing import NamedTuple, Tuple, List

import pyautogui
import pyperclip
from PIL import ImageGrab, Image, ImageChops
import pyperclip as pc #pip install pyperclip


def setup_logger(level=logging.DEBUG, disable_imported: bool = False) -> logging.Logger:
    import logging.config

    if disable_imported:
        logging.config.dictConfig({'version': 1, 'disable_existing_loggers': True})

    logger_ = logging.getLogger(__file__)

    stream_handler = logging.StreamHandler(sys.stdout)
    # noinspection SpellCheckingInspection
    stream_handler.setFormatter(logging.Formatter("[{asctime}] [{levelname:7}] {message}", style='{'))
    logger_.addHandler(stream_handler)

    filehandler = logging.FileHandler(f"{datetime.datetime.now():%Y-%m-%d}.log")
    # noinspection SpellCheckingInspection
    filehandler.setFormatter(logging.Formatter("[{asctime}] [{name}.{funcName}:{lineno}] [{levelname}] {message}",
                                               style='{'))
    logger_.addHandler(filehandler)

    logger_.setLevel(level)
    return logger_


def is_window_active(window_title: str = "Rec Room") -> bool:
    """
    Does not return before `window_title` becomes the active window
    Returns true when `window_title` becomes the active window

    :param window_title: The title of the window
    :return: When the window becomes active
    """
    if window_title not in (pyautogui.getActiveWindowTitle() or ""):  # getActiveWindowTitle is sometimes `None`
        print(f"Waiting for {window_title} to become the active window... ", end="\r", flush=True)
        # While RecRoom window is not active, sleep
        while window_title not in (pyautogui.getActiveWindowTitle() or ""):
            time.sleep(0.1)
        print(" " * 70, end="\r")  # Empty the last line in the console
        time.sleep(0.5)
    return True


class Colors(NamedTuple):
    text = (55, 57, 61)  # The color of text in the Variable Input field (black)
    white = (229, 225, 216)  # The white background of the Variable Input field
    green = (187, 205, 182)  # The Variable Input field sometimes turns green - this is that color.


class ImageCoords(NamedTuple):
    min_y: int
    min_x: int

    max_y: int
    max_x: int


def found_colors(main_color: tuple[int, int, int], coordinates: ImageCoords) -> bool:
    """
    Returns True if `main_color` is found in the given coordinates

    :param main_color: The color to compare the detected color to
    :param coordinates: Coordinates of the window of pixels to be checked and compared
    :return: If the color in any of the pixels match the `main_color`
    """

    def is_color(compare_color: tuple[int, int, int], main_color_: tuple[int, int, int], tolerance: int = 30) -> bool:
        """
        Compare `compare_color` to `main_color` with a given tolerance

        :param compare_color: The color that is being compared
        :param main_color_: The color that is being compared
        :param tolerance: How close the colors can be (1 - 255)
        :return: Is `compare_color` same/similar as `main_color`
        """
        return ((abs(compare_color[0] - main_color_[0]) < tolerance)
                and (abs(compare_color[1] - main_color_[1]) < tolerance)
                and (abs(compare_color[2] - main_color_[2]) < tolerance))

    image = ImageGrab.grab()

    for coords_x in range(coordinates.min_x, coordinates.max_x):
        if is_color(image.getpixel((coords_x, coordinates.min_y)), main_color):
            return True

    return False


def color_in_coords(image: Image, color: Tuple[int, int, int], coordinates: List[Tuple[int, int]],
                    tolerance: int = 30) -> bool:
    """
    Returns True if `main_color` is found in the given coordinates given a tolerance

    :param image: The image from which the colors to compare will be taken
    :param color: The color to compare the detected color to
    :param coordinates: Coordinates of the window of pixels to be checked and compared
    [(top_left_corner), (bottom_right_corner)]
    :param tolerance: Max variation between colors
    :return: If the color in any of the pixels match the `main_color`
    """

    # coordinates: [
    #   (min_x, min_y)
    #   (max_x, max_y)
    # ]

    def is_color(compare_color_: Tuple[int, int, int]) -> bool:
        """
        Compare `compare_color` to `main_color` with a given tolerance

        :param compare_color_: The color that is being compared to the `main_color`
        :return: Is `compare_color` same/+-tolerance as `main_color`
        """
        nonlocal color, tolerance
        return ((abs(compare_color_[0] - color[0]) < tolerance)
                and (abs(compare_color_[1] - color[1]) < tolerance)
                and (abs(compare_color_[2] - color[2]) < tolerance))

    image_colors = image.load()
    for y in range(coordinates[0][1], coordinates[1][1], 1):
        for x in range(coordinates[0][0], coordinates[1][0], 1):
            compare_color: Tuple[int, int, int] = image_colors[x, y]
            if is_color(compare_color_=compare_color):
                return True
    return False

VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'clear':0x0C,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'pause':0x13,
           'caps_lock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'page_up':0x21,
           'page_down':0x22,
           'end':0x23,
           'home':0x24,
           'left_arrow':0x25,
           'up_arrow':0x26,
           'right_arrow':0x27,
           'down_arrow':0x28,
           'select':0x29,
           'print':0x2A,
           'execute':0x2B,
           'print_screen':0x2C,
           'ins':0x2D,
           'del':0x2E,
           'help':0x2F,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'numpad_0':0x60,
           'numpad_1':0x61,
           'numpad_2':0x62,
           'numpad_3':0x63,
           'numpad_4':0x64,
           'numpad_5':0x65,
           'numpad_6':0x66,
           'numpad_7':0x67,
           'numpad_8':0x68,
           'numpad_9':0x69,
           'multiply_key':0x6A,
           'add_key':0x6B,
           'separator_key':0x6C,
           'subtract_key':0x6D,
           'decimal_key':0x6E,
           'divide_key':0x6F,
           'F1':0x70,
           'F2':0x71,
           'F3':0x72,
           'F4':0x73,
           'F5':0x74,
           'F6':0x75,
           'F7':0x76,
           'F8':0x77,
           'F9':0x78,
           'F10':0x79,
           'F11':0x7A,
           'F12':0x7B,
           'F13':0x7C,
           'F14':0x7D,
           'F15':0x7E,
           'F16':0x7F,
           'F17':0x80,
           'F18':0x81,
           'F19':0x82,
           'F20':0x83,
           'F21':0x84,
           'F22':0x85,
           'F23':0x86,
           'F24':0x87,
           'num_lock':0x90,
           'scroll_lock':0x91,
           'left_shift':0xA0,
           'right_shift ':0xA1,
           'left_control':0xA2,
           'right_control':0xA3,
           'left_menu':0xA4,
           'right_menu':0xA5,
           'browser_back':0xA6,
           'browser_forward':0xA7,
           'browser_refresh':0xA8,
           'browser_stop':0xA9,
           'browser_search':0xAA,
           'browser_favorites':0xAB,
           'browser_start_and_home':0xAC,
           'volume_mute':0xAD,
           'volume_Down':0xAE,
           'volume_up':0xAF,
           'next_track':0xB0,
           'previous_track':0xB1,
           'stop_media':0xB2,
           'play/pause_media':0xB3,
           'start_mail':0xB4,
           'select_media':0xB5,
           'start_application_1':0xB6,
           'start_application_2':0xB7,
           'attn_key':0xF6,
           'crsel_key':0xF7,
           'exsel_key':0xF8,
           'play_key':0xFA,
           'zoom_key':0xFB,
           'clear_key':0xFE,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,
           '`':0xC0}


def paste(x):
    getActiveWindow()
    pc.copy(x.rstrip('\r\n'))
    pyautogui.hotkey('ctrl','v')

def ctrlA():
    getActiveWindow()
    pyautogui.hotkey('ctrl','a')
            

def getActiveWindow(window_title: str = "Rec Room") -> bool: #From Reny
    if window_title not in (pyautogui.getActiveWindowTitle() or ""): #If Rec Room is not the main window
        print("Waiting for Rec Room to be the active window... ")
        while window_title not in (pyautogui.getActiveWindowTitle() or ""): #While Rec Room is not main window
            sleep(0.1)
        sleep(0.5)
    return(True)

def sleep(t):
    time.sleep(t)

def grabSquareAtPosition(name,pos):
    image = ImageGrab.grab()
    X,Y = pos
    square_A = 60
    half = square_A/2

    bbox = (X-half,Y-half,X+square_A,Y+square_A)
    sub_image = image.crop(bbox)
    sub_image.save("data/"+name+".png")



def images_are_equal(img1,img2):
    if img1.size != img2.size:
        print("wrong size")
        return False
    diff = ImageChops.difference(img1,img2)
    threshold = 50
    diff = diff.point(lambda x: 0 if x < threshold else 255)
    exceeded_points = sum(1 for pixel in diff.getdata() if pixel == 255)
    
    print(f"Points where threshold was exceeded: {exceeded_points}")
    return not diff.getbbox(), exceeded_points

def compareSquareAtPosition(name,pos):
    image = ImageGrab.grab()
    X,Y = pos
    square_A = 60
    half = square_A/2
    template_name = "data/"+name+".png"
    template = Image.open(template_name)

    bbox = (X-half,Y-half,X+square_A,Y+square_A)
    sub_image = image.crop(bbox)
    return images_are_equal(template,sub_image)

def waitForMenu(name,pos):
    for i in range(10):
        r,exceeded_points = compareSquareAtPosition(name,pos[name])
        if r:
            grabSquareAtPosition(name+"_success",pos[name])
            return True
        print("waiting for "+ name +" menu to appear")
        sleep(0.1)
    grabSquareAtPosition(name+"_fail",pos[name])
    return False

def waitForMenu2(name,pos):
    for i in range(10):
        r,exceeded_points = compareSquareAtPosition(name,pos[name])
        if r:
            grabSquareAtPosition(name+"_success",pos[name])
            return True, exceeded_points
        print("waiting for "+ name +" menu to appear")
        sleep(0.1)
    grabSquareAtPosition(name+"_fail",pos[name])
    return False, exceeded_points

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
    print("Data chunk size: ", chunk_size)
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
    getActiveWindow()
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

def dropMakerPen(delay=1):
    press("z")
    sleep(delay)

def makerPenMenu(delay=1):
    press("f")
    sleep(delay)
