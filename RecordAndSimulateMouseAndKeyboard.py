import PyHook3
import os
import win32api
import win32con
import time

global IsRecording
global InfoArr
global IsPlaying
global loop

InfoArr = []
IsRecording = False
IsPlaying = False
loop = 10
  
def OnMouseEvent(event):
##  print('MessageName:',event.MessageName)
##  print('Message:',event.Message)
##  print('Time:',event.Time)
##  print('Window:',event.Window)
##  print('WindowName:',event.WindowName)
##  print('Position:',event.Position)
##  print('Wheel:',event.Wheel)
##  print('Injected:',event.Injected)
##  print('---')
  SpecialMouseEvents(event)
  
  # return True to pass the event to other handlers
  # return False to stop the event from propagating
  return True

def OnKeyboardEvent(event):
##  print('MessageName:',event.MessageName)
##  print('Message:',event.Message)
##  print('Time:',event.Time)
##  print('Window:',event.Window)
##  print('WindowName:',event.WindowName)
##  print('Ascii:', event.Ascii, chr(event.Ascii))
##  print('Key:', event.Key)
##  print('KeyID:', event.KeyID)
##  print('ScanCode:', event.ScanCode)
##  print('Extended:', event.Extended)
##  print('Injected:', event.Injected)
##  print('Alt', event.Alt)
##  print('Transition', event.Transition)
##  print('---')
  SpecialKeyEvents(event)

  # return True to pass the event to other handlers
  # return False to stop the event from propagating
  return True

def SpecialKeyEvents(event):
  global IsRecording
  global InfoArr
  global IsPlaying
  if  "Escape" == event.Key:
      Exit()
  
  if IsPlaying == True:
      return   
  if "F10" == event.Key and IsRecording == True:
      IsRecording = False     
      stoprecord()
  if  IsRecording == True:
      InfoArr.append([event.MessageName,event.KeyID,event.Time])
  if "F9" == event.Key and IsRecording == False:
      IsRecording = True 
      startrecord()

  if  "F12" == event.Key:
      Play()
      
def SpecialMouseEvents(event):
    global IsRecording
    global InfoArr
    global IsPlaying
    if IsPlaying == True:
        return    
    if  IsRecording == True:
        InfoArr.append([event.MessageName,event.Position,event.Time])
        
def startrecord():
    print("Start Record")
def stoprecord():
    print("Stop Record")
def Exit():
    print("Exit")
    os._exit(0)
    
def Play():
    global InfoArr
    global IsPlaying
    global loop
    locloop = loop
    
    if IsPlaying == True:
      return
    if IsPlaying == False:
        IsPlaying = True
    t0 = InfoArr[0][2]
    for i in InfoArr:
      i[2] = i[2] - t0
      
#循环loop次数播放
    for i in range(loop - 1):
      for i in InfoArr:
        if i[0] == 'mouse left down':
          time.sleep(i[2]/1000)
          x , y = i[1]
          mouse_left_click(x, y, 1)
        if i[0] == 'key down':
          time.sleep(i[2]/1000)
          win32api.keybd_event(i[1], 0, 0, 0)
          win32api.keybd_event(i[1], 0, win32con.KEYEVENTF_KEYUP, 0)
    EndPlay()
##    print(InfoArr)
    
def EndPlay():
    global IsPlaying
    if IsPlaying == True:
        IsPlaying = False
    print("EndPlay")
    

##鼠标移动和鼠标左键输出
def mouse_move(new_x, new_y):
    if new_y is not None and new_x is not None:
        point = (new_x, new_y)
        win32api.SetCursorPos(point)
      
def mouse_left_click(new_x=None, new_y=None, times=1):
    """
    鼠标左击事件
    :param new_x: 新移动的坐标x轴坐标
    :param new_y: 新移动的坐标y轴坐标1506240215
    :param times: 点击次数
    """
    mouse_move(new_x, new_y)
    
    while times:
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        times -= 1
##分割线
    
def main():
    # create the hook mananger
    hm = PyHook3.HookManager()
    # register two callbacks
    hm.MouseAllButtonsDown = OnMouseEvent
    hm.KeyDown = OnKeyboardEvent

    # hook into the mouse and keyboard events
    hm.HookMouse()
    hm.HookKeyboard()
    if __name__ == '__main__':
      import pythoncom
      pythoncom.PumpMessages()
      
main()
