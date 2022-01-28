import pytesseract as tess
from PIL import Image
import pyautogui
import time
import keyboard
import winsound
import win32api
import win32con
import os


#globel variables
textList= []         
start = 1                #for main loop
shotCount = 0


x=70
y=66
boxHeight = 38
boxWidth = 386 #inside the line (not include the black line)
xAxis = 0
yAxis = 0


def findMousePoint(x,y):
    global xAxis,yAxis
    def findY(x,y):
        while 1:
            red = pyautogui.pixel(x,y)[0]
            #pyautogui.moveTo(x,y)            
            if red != 0:
                y += 1
            else:
                break
        y+=1
        return y

    yAxis = findY(x,y)
    def findX(x,y):
        while 1:
            red = pyautogui.pixel(x,y)[0]
            if red != 0:
                x -= 1
            else:
                break
        x += 1
        return x
    xAxis = findX(x,yAxis)
    shot()

# ScreenShot
def shot():
    
    winsound.PlaySound('C:\\Users\\anshe\\Desktop\\getinfo\\sound\\001.wav',0)
    global shotCount
    time.sleep(2)
    #1
    im = pyautogui.screenshot(region=(xAxis, yAxis, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1

    #2
    im = pyautogui.screenshot(region=(xAxis+boxWidth+1, yAxis, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #3
    im = pyautogui.screenshot(region=(xAxis+int(boxWidth*2)+3, yAxis, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #4
    im = pyautogui.screenshot(region=(xAxis, yAxis+boxHeight+1, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    #5
    im = pyautogui.screenshot(region=(xAxis+boxWidth+2, yAxis+boxHeight+1, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #6
    im = pyautogui.screenshot(region=(xAxis+(int(boxWidth*2))+3, yAxis+boxHeight+1, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #7
    im = pyautogui.screenshot(region=(xAxis, yAxis+(int(boxHeight*2))+5, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #8
    im = pyautogui.screenshot(region=(xAxis+boxWidth+2, yAxis+(int(boxHeight*2))+5, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #9
    im = pyautogui.screenshot(region=(xAxis+int(boxWidth*2)+3, yAxis+(int(boxHeight*2))+5, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #10
    im = pyautogui.screenshot(region=(xAxis, yAxis+(int(boxHeight*3))+7, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #11
    im = pyautogui.screenshot(region=(xAxis+boxWidth+2, yAxis+(int(boxHeight*3))+7, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #12
    im = pyautogui.screenshot(region=(xAxis+int(boxWidth*2)+3, yAxis+(int(boxHeight*3))+7, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #13
    im = pyautogui.screenshot(region=(xAxis, yAxis+(int(boxHeight*4))+17, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #14
    im = pyautogui.screenshot(region=(xAxis+boxWidth+2, yAxis+(int(boxHeight*4))+17, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #15
    im = pyautogui.screenshot(region=(xAxis+int(boxWidth*2)+3, yAxis+(int(boxHeight*4))+17, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
   
    #16
    im = pyautogui.screenshot(region=(xAxis, yAxis+(int(boxHeight*5))+19, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #17
    im = pyautogui.screenshot(region=(xAxis+boxWidth+2, yAxis+(int(boxHeight*5))+19, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount += 1
    
    #18
    im = pyautogui.screenshot(region=(xAxis+int(boxWidth*2)+3, yAxis+(int(boxHeight*5))+19, boxWidth, boxHeight))
    im.save(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
    shotCount = 0 
    print('ScreenShot Done ...........')
    time.sleep(1)
    translate()


def translate():    
    global textList,shotCount
    tess.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    while shotCount < 18:
        img = Image.open(f'C:\\Users\\anshe\\Desktop\\getinfo\\images\\image{shotCount:02}.png')
        img = img.resize(tuple(3*x for x in img.size))
        text = tess.image_to_string(img)
        textList.append(text.split('\n')[0])
        # textList.append((filter(None, text.splitlines())))
        shotCount += 1
    shotCount = 0
    print('Translation Completed ..........')
    mainKeyboard()

def mainKeyboard():
    win32api.SetCursorPos((350,405))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    time.sleep(1)
    for i,j in enumerate(textList, 1):
        print(i,'.  ',j)

    for data in textList:
        for char in data:
            time.sleep(0.1)
            if char.isupper():
                keyboard.press_and_release(f'shift+{char}')
            elif char.islower():
                keyboard.press_and_release(char)
            elif char == '£':
                keyboard.press_and_release('f')
            else:
                keyboard.press_and_release(char)

        keyboard.press_and_release('tab')
    winsound.PlaySound('C:\\Users\\anshe\\Desktop\\getinfo\\sound\\009.wav',0)
    textList.clear()
    print(f'DOne: {start}')
    keyboard.wait('esc')
    print('\n'*4)

print("Start...")

while start <= 70:
    # time.sleep(5)
    print("Going to SHot")
    findMousePoint(x,y)
    start += 1

