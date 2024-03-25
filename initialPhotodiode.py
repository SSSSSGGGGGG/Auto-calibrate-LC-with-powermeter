# -*- coding: utf-8 -*-
"""
Created on Wed Jan 31 14:20:13 2024

@author: Laboratorio
"""

import win32gui
import win32com.client
import time
import traceback

def iniPhotodiode (a,newrange):
   if a:
    try:
     OphirCOM = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement")     
     # Stop & Close all devices
     OphirCOM.StopAllStreams() 
     OphirCOM.CloseAll()
     # Scan for connected Devices
     DeviceList = OphirCOM.ScanUSB()     
     print(DeviceList) # device number of photodiode
     for Device in DeviceList:   	# if any device is connected
      DeviceHandle = OphirCOM.OpenUSBDevice(Device)	# open first device
      exists = OphirCOM.IsSensorExists(DeviceHandle, 0)
      if exists:
       print('\n----------Data for S/N {0} ---------------'.format(Device))
       ranges = OphirCOM.GetRanges(DeviceHandle, 0)
       print ("Built-in ranges:",ranges)
       mode=OphirCOM.GetMeasurementMode(DeviceHandle,0)       
       print('The mode of photodiode: ', mode)
       # set new range
       OphirCOM.SetRange(DeviceHandle, 0, newrange)
       testRange=OphirCOM.GetRanges(DeviceHandle, 0)[0]
       print('The selected range : ',testRange)
     
    except OSError as err:
     print("OS error: {0}".format(err))
    except:
     traceback.print_exc()
    return OphirCOM