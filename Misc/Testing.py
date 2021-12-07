# Testing Script
import System
import Control
import UI
import Guide

windowHandle = 0

def Init():
	global windowHandle
	UI.SetScriptName("Testing")
	windowHandle = System.FindWindow("MyPythonWindowClass", "win32gui test")
	Guide.AddEventHandler("Event_HandleCreateComponent","Testing")
	Guide.AddEventHandler("Event_ComponentItemSelected","Testing")
	return 0
                
def Event_HandleCreateComponent(szComponentId):
	global windowHandle
	if windowHandle != 0:
		System.SendText(windowHandle, 1, szComponentId)
	return 0

def Event_ComponentItemSelected(ItemId):
	global windowHandle
	if windowHandle != 0:
		displyText = UI.GetComponentItemDisplayText(ItemId)
		System.SendText(windowHandle, 2, displyText)
	return 0
