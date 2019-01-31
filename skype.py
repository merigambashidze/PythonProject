# -*- coding: utf-8 -*-
import Skype4Py
from win32api import *
from win32gui import *
import win32con
import sys, os
import struct
import time
import re
import urllib2

# Create an instance of the Skype class.
skype = Skype4Py.Skype()

# Connect the Skype object to the Skype client.
skype.Attach()

# Obtain some information from the client and print it out.

def message_check(Notification):
    noti = str(Notification)
    leng =len(noti)
    received = noti[leng-8:]
    if received=="RECEIVED":
        notification = str(Notification)
        noti1 = notification[12:]
        length_noti1 = len(noti1)-16
        noti2 = noti1[:length_noti1]
        received_message =skype.Message(Id = noti2).Body
        #here goes main code
        pattern = "//www.youtube.com"
        if re.findall(pattern, received_message):
            link_list = received_message.split(" ")
            index = 0
            for n in link_list:
                youtube = link_list[index]
                if link_list[index][:23] == "https://www.youtube.com" or link_list[index][:22]=="http://www.youtube.com":
                    if len(link_list[index]) > 42:
                        #parsing url
                        youtube_link = link_list[index]
                        response = urllib2.urlopen(youtube_link)
                        html = response.read()
                        pattern1 = "<title>"
                        if re.findall(pattern1,html):
                            new_html = html.split("<title>")
                            new_new_html  = new_html[1].split("</title>")
                            song_name =  new_new_html[0]
                            balloon_tip("Youtube", song_name)


                    
                index = index +1

            
            balloon_tip("Youtube", song_name)

skype.RegisterEventHandler('Notify', message_check)

class WindowsBalloonTip:
    def __init__(self, title, msg):
        message_map = { win32con.WM_DESTROY: self.OnDestroy,}

        # Register the window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = 'PythonTaskbar'
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = RegisterClass(wc)

        # Create the window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow(classAtom, "Taskbar", style, 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        # Icons managment
        iconPathName = os.path.abspath(os.path.join( sys.path[0], 'balloontip.ico' ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            hicon = LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, 'Tooltip')

        # Notify
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20, hicon, 'Balloon Tooltip', msg, 200, title))
        # self.show_balloon(title, msg)
        time.sleep(5)

        # Destroy
        DestroyWindow(self.hwnd)
        classAtom = UnregisterClass(classAtom, hinst)
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

# Function
def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)
