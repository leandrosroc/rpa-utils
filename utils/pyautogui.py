import pyautogui
import time

def click(x, y, clicks=1, interval=0.0, button='left'):
    pyautogui.click(x=x, y=y, clicks=clicks, interval=interval, button=button)

def write(text, interval=0.05):
    pyautogui.write(text, interval=interval)

def press(key):
    pyautogui.press(key)

def hotkey(*args):
    pyautogui.hotkey(*args)

def screenshot(path):
    pyautogui.screenshot(path)

def locate_on_screen(image, confidence=0.8):
    while True:
        try:
            pos = pyautogui.locateCenterOnScreen(image, confidence=confidence)
            if pos:
                return pos
        except Exception:
            pass
        time.sleep(0.5)

def move_to(x, y, duration=0.25):
    pyautogui.moveTo(x, y, duration=duration)