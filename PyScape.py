import os
from os import walk
import random
import time as t
import win32api
import win32com.client
import win32con
import win32gui

import cv2 as cv
import numpy as np
from PIL import ImageGrab
from pynput.keyboard import Listener as keyL
from pynput.mouse import Listener as mouseL


def winFix(clientwindow):
    win32gui.ShowWindow(clientwindow, win32con.SW_SHOWNORMAL)
    win32gui.MoveWindow(clientwindow, 0, 0, 775, 535, True)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)


def oplisten():
    listener2 = mouseL(on_click=on_click)
    listener2.start()
    listener2.join()


while True:
    try:
        opflop = str(input("Enable the autoclicker? (y, n) : "))
        if not opflop.isalpha():
            print("Yes or No doesn't have 1's or 0's.")
            continue
        if opflop[0].lower() == "y":
            oplisten()
            print("Autoclicker enabled.")
        if opflop[0].lower() == "n":
            print("Autoclicker disabled.")
        break
    except ValueError:
        print("[Y]es, [N]o")
        continue


print("OSRS Object Detection and Autoclicker")
print("press 'Numpad 8' to start Object Detection")
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
# Vars
tb = []  # click coordinates
td = []  # time delays to set each click
tds = []  # Saves state of the previous time.
oorc = []
mlogs = 0
oocc = {}
delay = {}  # Time delays in dictionary form
countclick = 0
counters = 0


def move(x, y):
    win32api.SetCursorPos((x, y))


def click(x, y):
    win32api.SetCursorPos((x, y))
    t.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    t.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def vidcapswitch():
    global oocc
    global countclick
    global counters
    global mlogs
    vidcap = True
    while vidcap:
        tlogs = 0
        ss = np.array(ImageGrab.grab(bbox=(0, 0, 775, 535)))
        # cv.putText(ss, str(fps), (25, 30), cv.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255))
        ssg = cv.cvtColor(ss, cv.COLOR_BGR2GRAY)
        filepath = next(walk(r'imageres\alll'), (None, None, []))[2]
        for imgsrc in enumerate(filepath):
            targpath = "C:\\PyProjects\\PyScape\\PyScape\\imageres\\alll\\" + imgsrc[1]
            targ = cv.imread(str(targpath), 0)
            print(targpath)
            print(targ)
            w, h = targ.shape[::-1]
            res = cv.matchTemplate(ssg, targ, cv.TM_CCOEFF_NORMED)
            thresh = 0.80
            loc = np.where(res >= thresh)
            for pt in zip(*loc[::-1]):
                cv.rectangle(ss, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 1)
                # ooll = [pt[0], pt[1], pt[0] + w, pt[1] + h]
                # oocc[imgsrc[1]] = ooll
                # print(oocc)
                if filepath == "1513":
                    tlogs += 1
            if tlogs is not mlogs:
                mlogs = tlogs
                print(f"Mage Logs: {mlogs}")
        cv.imwrite('res.png', ss)
        cv.putText(ss, f"Mage Logs: {mlogs}", (580, 480), cv.FONT_HERSHEY_SIMPLEX, 0.7, (120, 50, 120), 2)
        cv.imshow('Big Brother', ss)
        if cv.waitKey(25) & 0xFF == ord('q'):
            cv.destroyAllWindows()
            vidcap = False


"""         
                elif ls in invT:
                    for mageGone in enumerate(invT):
                        mgg = mageGone[1]
                        mg = np.array(ImageGrab.grab(bbox=(mgg[0], mgg[1], mgg[2], mgg[3])))
                        mmg = cv.cvtColor(mg, cv.COLOR_BGR2GRAY)
                        matched = cv.matchTemplate(mmg, targ, cv.TM_CCOEFF_NORMED)
                        if matched.item(0, 0) <= 0.10:
                            try:
                                if mgg in invT:
                                    invT.remove(ls)
                                    inv -= 1
                                    print("Mage Log Gone")
                                continue
                            except ValueError as pp:
                                print(pp, "Errrrrr")
                                invT.pop(len(invT)-1)
                if countclick >= 10:
                    x = (ls[2] + ls[0]) / 2
                    x = int(x)
                    y = (ls[3] + ls[1]) / 2
                    y = int(y)
                    # click(x, y)
                    countclick = 0
                    print(f'Sent 1 Click to X: {x}, Y:{y} : Found"{imgsrc[1]}"')
            countclick += 1

        """


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
        for o in range(0, 60000):
            ll = random.randint(20, 55)
            if random.randint(0, 10000) > 9900:  # this is a random chance of standing at bank
                print(f"script waiting for: {ll} seconds")
                for j in range(0, ll - 1):
                    xr = random.randint(0, 1920)
                    yr = random.randint(0, 1080)
                    print(f"Moving mouse to (X: {xr}, Y:{yr})")
                    move(xr, yr)
                    t.sleep(1)
            x3 = 0
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
                        print(f"Moving mouse to (X: {xr}, Y: {yr})")
                        move(xr, yr)
                        t.sleep(1)
                else:
                    t.sleep(sl)


def login():
    print("Not Implemented.")


def close():
    print("ESC Pressed, Closing Script")
    os._exit(0)


def on_press(key):
    hotkeys = {
        '<105>': login,
        '<104>': vidcapswitch,
        'Key.esc': close
    }
    if hotkeys.get(str(key)) in hotkeys.values():
        hotkeys.get(str(key))()
    else:
        print(key)
        print("Unknown HK Detected")


listener1 = keyL(on_press=on_press)
listener1.start()
listener1.join()