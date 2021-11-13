import pyautogui, sys
import time

try:
    while True:
        x, y = pyautogui.position()
        print(f'X: {x}')
        print(f'Y: {y}')
        time.sleep(1)

except:
    pass
