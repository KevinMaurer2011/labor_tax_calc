import pyautogui
import time


#contact
pyautogui.moveTo(52, 170, duration=0.5)
pyautogui.click()
#estimator
pyautogui.moveTo(2477, 287, duration=0.5)
pyautogui.click()
#jason
pyautogui.moveTo(2473, 415, duration=0.5)
pyautogui.click()
#inspection tab
pyautogui.moveTo(185, 170, duration=0.5)
pyautogui.click()
#open date picker
pyautogui.moveTo(1477, 259, duration=0.05)
pyautogui.click()
#select today
pyautogui.moveTo(1591, 407, duration=0.5)
pyautogui.click()
#site type dropdown
pyautogui.moveTo(1558, 286, duration=0.5)
pyautogui.click()
#virtual
pyautogui.moveTo(1555, 403, duration=0.5)
pyautogui.click()
#settlement tab
pyautogui.moveTo(493, 171, duration=0.5)
pyautogui.click()
#open settlement decision
pyautogui.moveTo(2168, 473, duration=0.5)
pyautogui.click()
#set repairable
pyautogui.moveTo(2168, 506, duration=0.5)
pyautogui.click()
#save file
pyautogui.moveTo(30, 85, duration=0.5)
pyautogui.click()
time.sleep(5)
#convert to job
pyautogui.moveTo(296, 91, duration=0.5)
pyautogui.click()
time.sleep(5)
#job number box
pyautogui.moveTo(1306, 690, duration=0.5)
pyautogui.click()
pyautogui.typewrite('Test')
