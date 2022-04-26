import os
from os import walk
import random
import time as t
import win32api
import win32com.client
import win32con
import win32gui

import cv2 as cv
import keyboard as k
import numpy as np
import pytesseract
from PIL import ImageGrab
from matplotlib import pyplot as plt
from pynput.keyboard import Listener as keyL
from pynput.mouse import Listener as mouseL

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Pexo\AppData\Local\tesseract\Tesseract-OCR\tesseract.exe'


def winFix(clientwindow):
    win32gui.ShowWindow(clientwindow, win32con.SW_SHOWNORMAL)
    win32gui.MoveWindow(clientwindow, 0, 0, 775, 535, True)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)


print("Mouse Event Python Auto Clicker")
print("Script is designated to use the side buttons of your mouse.")
print("Press 'Numpad 9' to auto-login")
print("Press 'ESC' to close program.")

while True:
    try:
        task = int(input("Enter delay in seconds after each loop: "))
        break
    except ValueError:
        print("Please input integer only...")
        continue
while True:
    try:
        client = str(input("Enter Runescape Client Name: "))
        hwnd = win32gui.FindWindow(None, client)
        if hwnd == 0:
            print("Client Name Not Found, check spelling")
            continue
        winFix(hwnd)
        break
    except BaseException as e:
        print("Client Name Not Found, check spelling")
        print(e)
        continue
tb = []  # click coordinates
td = []  # time delays to set each click
tds = []  # Saves state of the previous time.
delay = {}  # Time delays in dictionary form


def move(x, y):
    win32api.SetCursorPos((x, y))


def click(x, y):
    win32api.SetCursorPos((x, y))
    t.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    t.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def vidcapswitch():
    vidcap = True
    while vidcap:
        ss = np.array(ImageGrab.grab(bbox=(0, 0, 775, 535)))
        # fps = cv.getTickFrequency() / (cv.getTickCount() - timer)
        # cv.putText(ss, str(fps), (25, 30), cv.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255))
        ssg = cv.cvtColor(ss, cv.COLOR_BGR2GRAY)
        filepath = next(walk(r'imageres\all'), (None, None, []))[2]
        for imgsrc in enumerate(filepath):
            targpath = "C:\\PyProjects\\PyScape\\PyScape\\imageres\\all\\" + imgsrc[1]
            targ = cv.imread(str(targpath), 0)
            w, h = targ.shape[::-1]
            res = cv.matchTemplate(ssg, targ, cv.TM_CCOEFF_NORMED)
            thresh = 0.80
            loc = np.where(res >= thresh)
            for pt in zip(*loc[::-1]):
                cv.rectangle(ss, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 1)
                print(pt[0], pt[1])
            cv.imwrite('res.png', ss)
        cv.imshow('Runelite Source', ss)
        if cv.waitKey(25) & 0xFF == ord('q'):
            cv.destroyAllWindows()
            vidcap = False


def on_click(x, y, button, pressed):
    trt = len(tb) > 0
    if pressed is True and button is button.x1:  # this is your back side button
        tb.append(x)
        tb.append(y)
        if len(tb) > 0:
            timedelay = t.time_ns() / 1000000000
            td.append(timedelay)
        if len(td) > 1:
            tds.clear()  # we are clearing the last delays set.
            dll = td[1] - td[0]  # calculating current click to last click time.
            dl = random.uniform(dll, (dll * 1.1))  # this is the delay between each click
            delay[len(delay)] = dl  # this is just tracking the delays set.
            tds.append(td[1])  # holds the last value of the current click time
            print(f"Adding Click at X: {x}, Y: {y}, with a {dl} second delay")
            td.clear()  # clears the current time delay calcs.
            td.append(tds[0])  # sets last click time that was tracked.

    if pressed is True and trt is True and button is button.x2:  # Front Side will start script
        print("script is starting in 5 seconds")
        t.sleep(5)
        print("TB: ", tb)
        print("TD: ", td)
        print("TDS: ", tds)
        print("Delay: ", delay)
        for o in range(0, 60000):
            ll = random.randint(20, 55)
            if random.randint(0, 10000) > 9900:  # this is a random chance of standing at bank
                print("script waiting for:", ll, "seconds")
                for j in range(0, ll - 1):
                    xr = random.randint(0, 1920)
                    yr = random.randint(0, 1080)
                    print("Moving mouse to (", xr, ",", yr, ")")
                    move(xr, yr)
                    t.sleep(1)
            x3 = 0
            x4 = 0
            for x4 in range(0, len(delay) + 1):
                sl = delay.get(x4)
                xx = tb[x3] + random.randint(-4, 4)
                x3 += 1
                yy = tb[x3] + random.randint(-4, 4)
                x3 += 1
                click(xx, yy)
                print(f"Clicked X: {xx}, Y: {yy}")
                if sl is None:
                    sl = task
                    for k in range(0, sl - 1):
                        xr = random.randint(0, 1920)
                        yr = random.randint(0, 1080)
                        print("Moving mouse to (", xr, ",", yr, ")")
                        move(xr, yr)
                        t.sleep(1)
                else:
                    t.sleep(sl)


def login():
    print("Logging in.")
    t.sleep(1)
    click(466, 320)
    k.write("", 0.1)  # edit password here
    t.sleep(1)
    click(305, 349)
    t.sleep(7)
    click(390, 359)
    t.sleep(1)
    click(732, 17)
    print("Done.")


def close():
    print("ESC Pressed, Closing Script")
    os._exit(0)


hotkeys = {
    '<105>': login,
    '<104>': vidcapswitch,
    'Key.esc': close
}


def on_press(key):
    if hotkeys.get(str(key)) in hotkeys.values():
        hotkeys.get(str(key))()
    else:
        print("Unknown HK Detected")


listener1 = mouseL(on_click=on_click)
listener2 = keyL(on_press=on_press)
listener1.start()
listener2.start()
listener1.join()
listener2.join()