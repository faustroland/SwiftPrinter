import win32api
import win32con
import pyautogui
import time
import ctypes

from common import *

def CAPSLOCK_STATE():
    
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_CAPITAL = 0x14
    return hllDll.GetKeyState(VK_CAPITAL)

def get_capslock_state():
    import ctypes
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_CAPITAL = 0x14
    return hllDll.GetKeyState(VK_CAPITAL)

def IS_CAPSON():
    CAPSLOCK = CAPSLOCK_STATE()
    if ((CAPSLOCK) & 0xffff) != 0:
        return 1
    else:
        return 0
def kp(key,length=0.1,wait=0.1):
    #getActiveWindow()
    win32api.keybd_event(key,0,0,0)  
    time.sleep(length)
    win32api.keybd_event (key,0, win32con.KEYEVENTF_KEYUP, 0) # key is released    
    if wait:
        time.sleep(wait)
def CAPSOFF():
    if get_capslock_state()==1:
        kp(0x14,0.01,0.01)

def load_settings(file_path):
    settings = {}
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line by '='
            key, value = line.strip().split('=')
            # Store key-value pair in dictionary
            settings[key.strip()] = tuple(map(int, value.strip().split(',')))
    return settings

def save_settings(settings, file_path):
    with open(file_path, 'w') as file:
        for key, value in settings.items():
            # Convert value tuple to comma-separated string
            value_str = ','.join(map(str, value))
            # Write key-value pair to file
            file.write(f"{key}={value_str}\n")



def wait_until_capslock_on():
    while True:
        if IS_CAPSON():
            time.sleep(0.1)
            CAPSOFF()
            return pyautogui.position()
        time.sleep(0.1)


positions = load_settings("positions.txt")
def print_options(options):
    # Print options in multiple columns
    for i in range(0, len(options),3):
        if len(options[i:i+3])==3:
            print("{:<30}{:<30}{:<}".format(*options[i:i+3]))
        if len(options[i:i+3])==2:
            print("{:<30}{:<}".format(*options[i:i+2]))
        if len(options[i:i+3])==1:
            print("{:<}".format(*options[i:i+3]))


def print_welcome_message(settings):
    print("Choose which position you want to update:")
    options = ["ALL(0)"]
    options.extend([f"{name}({index + 1})" for index, (name, _) in enumerate(settings.items())])
    print_options(options)

def update_settings_loop(settings):
    while True:
        print_welcome_message(settings)
        choice = input("Enter your choice (0 to update all, or  X to save and exit): ")
        if choice.isdigit(): 
            choice = int(choice)
            if choice == 0:
                # Update all positions
                new_positions = {}
                for position_name, _ in settings.items():
                    print(f"\nMove the mouse cursor to position --> {position_name} <-- and turn on Caps Lock.")
                    new_position = wait_until_capslock_on()
                    new_positions[position_name] = new_position
                    print(f"Position {position_name} saved:", new_position)
                    print("\n\n")
                settings.update(new_positions)
                print("All positions updated successfully!")
            elif 1 <= choice <= len(settings):
                position_name = list(settings.keys())[choice - 1]
                print(f"Move the mouse cursor to position --> {position_name} <-- and turn on Caps Lock.")
                new_position = wait_until_capslock_on()
                kp(0x14,0.01,0.01)
                settings[position_name] = new_position
                win32api.SetCursorPos((1,1))
                sleep(1)
                grabSquareAtPosition(position_name,new_position)
                print(f"Position {position_name} updated successfully:", new_position)
            else:
                print("Invalid choice. Please enter a valid option.")
        elif choice.lower() == "x":
            print("Saving settings and exiting...")
            save_settings(settings, "positions.txt")
            print("Settings saved successfully. Exiting...")
            break
        else:
            print("Invalid input. Please enter a number or 'Save and Exit'.")

# Example usage:
if __name__ == "__main__":
    update_settings_loop(positions)
