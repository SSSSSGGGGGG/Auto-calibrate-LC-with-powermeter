# -*- coding: utf-8 -*-
"""
Created on Wed Jan 31 14:54:22 2024

@author: Laboratorio
"""

#import win32gui
import win32com.client


def close(end):
    if end=="y":
        OphirCOM = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement") 
        #win32gui.MessageBox(0, 'finished', '', 0)
        # Stop & Close all devices
        OphirCOM.StopAllStreams()
        OphirCOM.CloseAll()
        # Release the object
        OphirCOM = None
    else:
        print("continue your measurement")