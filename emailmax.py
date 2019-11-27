#https://stackoverflow.com/questions/2090464/python-window-activation

import pyautogui
import random
import pywinauto
import win32gui
import time
import wx
from pywinauto import application

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst

def print_running_applications():
	appwindows = get_app_list()
	for i in appwindows:
		print(i)


def select_subject_line():
#pyautogui.locateOnScreen('submit.png')
    subject_line_cood = pyautogui.locateCenterOnScreen('C:\\Ed Programmes\\Phyton\\images\\SubjectLineOutlook.png')
    print(subject_line_cood)
    pyautogui.moveTo(subject_line_cood, duration=0.2)
    pyautogui.click(subject_line_cood, button='left')

def move_to_application():
#pyautogui.locateOnScreen('submit.png')
    #location = pyautogui.locateCenterOnScreen('C:\\Ed Programmes\\Phyton\\images\\EmailSpimpiler.png')
    #print(location)
    pyautogui.moveTo(350,350, duration=0.2)
    
    
def send_email(name, email_address, subject_line):
	win32gui.ShowWindow(win32gui.FindWindow(None, "Inbox - ed@CourtHouseClinics.com - Outlook"),3)
	win32gui.SetForegroundWindow(win32gui.FindWindow(None, "Inbox - ed@CourtHouseClinics.com - Outlook"))
	pyautogui.moveTo(20, 87, duration=0.25)
	pyautogui.click(20, 87, button='left')
	print(pyautogui.position())
	pyautogui.typewrite(email_address + "\t\t\t\t" + subject_line + "\tHi " + name + "\n\nCheers\n\nEd")



class MyFrame(wx.Frame):    
    def __init__(self):
    	# create window
        super().__init__(parent=None, title='Email Simplifier', pos = (300,300))
        #save window to variable
        panel = wx.Panel(self)
        #Save the orientation to a variable        
        my_sizer = wx.BoxSizer(wx.VERTICAL)
        #Saves Reference to the text Box        
        self.text_ctrl = wx.TextCtrl(panel)
        #Add the text box to the widget
        my_sizer.Add(self.text_ctrl, 0, wx.ALL | wx.EXPAND, 5)    
        
        #Create Horitontal box
        top_buttons_box = wx.BoxSizer(wx.HORIZONTAL)
        #Add Horitltal box
        my_sizer.Add(top_buttons_box)  


        #Creates button (Email Max)   
        my_btn = wx.Button(panel, label='Email Max')
        #Bind the button to a event to fire when pushed

        my_btn.Bind(wx.EVT_BUTTON, self.on_press(email="Max@courthouseClinics.com",name="Max"))
        #Add the button box to the widget
        #my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)
        top_buttons_box.AddSpacer(30)
        top_buttons_box.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)

        #Creates button (Email Flavio)  
        my_btnFlavio = wx.Button(panel, label='Email Flavio') 
        #Bind the button to a event to fire when pushed
        my_btnFlavio.Bind(wx.EVT_BUTTON, self.on_press_Flavio)
        #Add the button box to the widget
        #my_sizer.Add(my_btnFlavio, 0, wx.ALL | wx.CENTER, 5)
        top_buttons_box.Add(my_btnFlavio, 0, wx.ALL | wx.CENTER, 5)



        panel.SetSizer(my_sizer)        
        self.Show()

        #Move cursor to application
        move_to_application()


    def on_press(self, event, email,name):
        value = self.text_ctrl.GetValue()
        send_email(name, email, value)
    


#        value = self.text_ctrl.GetValue()
#        if not value:
#            print("You didn't enter anything!")
#        else:
#            print(f'You typed: "{value}"')

    def on_press_Flavio(self, event, email, name):
        value = self.text_ctrl.GetValue()
        send_email("Flavio","flaviom@ABCLasers.co.uk",value)
    

if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()



#app = application.Application()
#app.start("Notepad.exe")

#
#import win32gui

#Awesome programme to get list of running programmes


print_running_applications()

print("Hello")
#window = win32gui.FindWindow(None, "Inbox - ed@CourtHouseClinics.com - Outlook")

#win32gui.ShowWindow(window,3)

#Maximise
#win32gui.ShowWindow(window,3)
#win32gui.SetForegroundWindow(window)
#print(window)


#Subject 754, 405
#TopMessage (524, 493)
#pyautogui.moveTo(754, 405, duration=0.25)
#pyautogui.click(754, 405, button='left')
#select_subject_line()
