import pyautogui
import logging
import os
import time
import ctypes
from PIL import Image
import screeninfo
import imagehash



file_path = r"C:\Users\AlejandroVelez\Documents\__PROCESS_WORK__\OpenBots -PBI -EOT\PBI\OPENBOTS.pbip"



logging.basicConfig(
    filename = "logs.log",
    encoding= "utf-8",
    level = logging.DEBUG,
    format = "%(asctime)s %(message)s"
)



def compare_images(cropped_path, reference_path, threshold = 10):

    image_ = Image.open(cropped_path)
    reference = Image.open(reference_path)
    hash_image = imagehash.average_hash(image_)
    hash_reference = imagehash.average_hash(reference)
    hamming_distance = hash_reference - hash_image
    
    similar = hamming_distance < threshold

    return similar, hamming_distance

# Launch pbix file Function
def launch_power_bi(file_path):
    try:
        os.startfile(file_path)
        logging.info("Launched Report Succesfully")
    except Exception as e:
        print(f"Error launching Power BI: {e}")
        logging.error(f"Error Did not launch .pbip File : {e}")


def maximize_power_bi_window():
    try:
        user = ctypes.WinDLL("user32")
        SW_MAXIMIZE = 3
        window = user.GetForegroundWindow()
        user.ShowWindow(window,SW_MAXIMIZE)
        logging.info("Maximized Power Bi Window")
    except Exception as e:
        print(f"Error Maximizing Window")


def refresh_pbi():
    try:

        screen = screeninfo.get_monitors()[0]
        centerx = screen.width //2
        centery = screen.height //2
        pyautogui.moveTo(
            x = centerx,
            y = centery,
             duration=1 )
        pyautogui.leftClick()
        time.sleep(1)
        pyautogui.keyDown("alt")
        time.sleep(1)
        pyautogui.keyUp("alt")
        time.sleep(1)
        pyautogui.press("H")
        time.sleep(1)
        pyautogui.press("R")
        logging.info("Hit refresh Button Succesfully")
    except Exception as e:
        logging.error(f"Did not hit the refresh Button Succesfully: {e}")




def wait_while_refresh():
    try:
        while True :
            time.sleep(15)
            #Take Screenshot
            screenshot = pyautogui.screenshot()
            screenshot.save("Screenshot_Raw.png")
            #Crop Image
            im = Image.open("Screenshot_Raw.png")
            w,h= im.size
            left = w/4 + w/9 
            right = left *1.75
            top = h/3
            bottom = top * 1.91
            im1 = im.crop((left,top,right,bottom))
            im1.save("Cropped.png")
            similar , distance = compare_images(
                cropped_path = "Cropped.png",
                reference_path="Reference.png"
            )
            print(f"{similar} - {distance}")

            if similar == False:
                break

            time.sleep(2)
    except Exception as e:
        logging(e)



def launch_file_options():
    try:
        screen = screeninfo.get_monitors()[0]
        centerx = screen.width //2 + 2
        centery = screen.height //2 + 2
        pyautogui.moveTo(
            x = centerx,
            y = centery,
             duration=1 )
        pyautogui.leftClick()
        pyautogui.leftClick()
        pyautogui.press("esc")
        time.sleep(1)
        pyautogui.keyDown("alt")
        time.sleep(1)
        pyautogui.keyUp("alt")
        time.sleep(1)
        pyautogui.press("f")
        time.sleep(1)
        pyautogui.press("enter")
        for i in range(8):
            time.sleep(0.5)
            pyautogui.press("down")
        for i in range(4):
            time.sleep(1)
            pyautogui.press("enter")
        time.sleep(3)
        pyautogui.typewrite("Automations_")
        pyautogui.press("enter")
        for i in range(2):
            pyautogui.press("tab")
        pyautogui.press("enter")

        while True:
            time.sleep(7)
            screenshot = pyautogui.screenshot()
            screenshot.save("Screenshot_Raw_2.png")
            im = Image.open("Screenshot_Raw_2.png")
            w,h= im.size
            left = w/4 + w/9 
            right = left *1.75
            top = h/3
            bottom = top * 1.81
            im1 = im.crop((left,top,right,bottom))
            im1.save("Cropped_2.png")
            
            similar , distance = compare_images(
                cropped_path = "Cropped_2.png",
                reference_path="Reference_2.png"
            )
            print(f"{similar} - {distance}")

            if similar == False:
                break

            pyautogui.press("enter")        
            time.sleep(10)
            pyautogui.press("tab")
            pyautogui.press("enter")
            
        
        logging.info("Laucnhed options side bar succesfully")
    except Exception as e:
        logging.error(f"Failed to launch options side bar: {e}")



def close_pbi():
    os.system(f"TASKKILL /F /IM PBIDesktop.exe")



# Launch pbix file Execution
launch_power_bi(file_path)
time.sleep(30)
maximize_power_bi_window()
time.sleep(5)
refresh_pbi()
wait_while_refresh()
time.sleep(5)
launch_file_options()
time.sleep(1)
close_pbi()



