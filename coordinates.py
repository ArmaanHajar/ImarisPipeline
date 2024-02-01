import pyautogui as auto
import keyboard as key
import time

stop = 0

while True:
    start = time.time()
    if key.is_pressed('z') and (start-stop) > 1:
        pos = auto.position()
        print(pos)
        stop = time.time()