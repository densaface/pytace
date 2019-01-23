#!/usr/bin/env python
# -*- coding: utf-8 -*-

# для отладки
# import ipdb; ipdb.set_trace()

import os
import libs.images as Img
import libs.ace_commander as ace_com
import time
import random
import win32gui, win32com.client

ace = ace_com.AceComs()
ace.log_file = r'c:\AutoClickExtreme\ace.log'
imlib = Img.imLib()
imlib.set_log(ace.log_file)
wins = ace_com.WindowMgr()

test_load = False # type True for activate test
if test_load:
    res_load = ace.loadAip(r'c:\AutoClickExtreme\aips\demo_notepad.aip')
    if not res_load:
        raise Exception("Error loading aip file. Check if it exists")
    print ("Success test load!")
    exit(0)

test_replay_only = False # test only start replay current aip script (without loading and
if test_replay_only:
    res_play = ace.Replay()
    if not res_play[0]:
        raise Exception("Error start Replay aip file.")
    print ("Success test start Replay!")
    exit(0)

test_play = False # test full aip replay with waiting end of Replaying
if test_play:
    res_play = ace.Replay(r'c:\AutoClickExtreme\aips\demo_notepad.aip')
    if not res_play[0]:
        raise Exception("Error playing aip file. Check if it exists")
    print ("Success test playing aip file!")
    exit(0)

test_stop = False
if test_stop:
    res_stop = ace.Replay()
    if not res_stop:
        raise Exception("Error stop Replay.")
    print ("Success stop Replay!")
    exit(0)

test_focus = False # test external lib for working with windows
if test_focus:
    w = ace.WindowMgr()
#   example with wildcard
#    w.find_window_wildcard("*Chrome*")
    w.find_window("Qt5QWindowIcon", "Linken Sphere")

#   workanound for python output error
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(w._handle)
    exit(0)

test_find_pict_via_ace = False
if test_find_pict_via_ace:
    brand = {'img': "./pics/brand.png", 'prefer_x': 10, 'prefer_y': 100,
             'max_radius': 100, 'blur': 1, 'deviation': 50}
    img_fn = "c:/AutoClickExtreme/screenshots/scr_1545921974.png"
#        ace.SaveDesktop(img_fn, 0, 0, 1024, 768)
    res_search = ace.GetImageCoord(img_fn, brand, 'c:/AutoClickExtreme')
    exit(0)

test_find_pict_via_python = False
if test_find_pict_via_python:
    img_fn = "c:/AutoClickExtreme/screenshots/scr_%d.png" % int(time.time())
    self.ace.SaveDesktop(img_fn, 0, 0, 1024, 768)
    res = imlib.find_image(img_fn, "template1", 0.8)
    if res:
        self.click(res[0] + 4 + random.randint(0, 4),
                   res[1] + 4 + random.randint(0, 4))
