import mss
import numpy as np
import cv2
import keyboard
import torch 
import serial
import time
import pygetwindow
import win32api
# Load the YOLOv5 model
model = torch.hub.load('ultralytics/yolov5', 'custom', path='C:/Users/Mr.Duy/OneDrive/Desktop/yolov5-master/yolov5-master/best.pt')

# Define screen size
ScreenSizeX = 600
ScreenSizeY = 600

ampX = 1
ampY = 1
# Selecting the correct game window
try:
    videoGameWindows = pygetwindow.getAllWindows()
    print("=== All Windows ===")
    for index, window in enumerate(videoGameWindows):
        # only output the window if it has a meaningful title
        if window.title != "":
            print("[{}]: {}".format(index, window.title))
    # have the user select the window they want
    try:
        userInput = int(input(
            "Please enter the number corresponding to the window you'd like to select: "))
    except ValueError:
        print("You didn't enter a valid number. Please try again.")
        exit()
    # "save" that window as the chosen window for the rest of the script
    videoGameWindow = videoGameWindows[userInput]
except Exception as e:
    print("Failed to select game window: {}".format(e))
    exit()

# Activate that Window
activationRetries = 30
activationSuccess = False
while (activationRetries > 0):
    try:
        videoGameWindow.activate()
        activationSuccess = True
        break
    except pygetwindow.PyGetWindowException as we:
        print("Failed to activate game window: {}".format(str(we)))
        print("Trying again... (you should switch to the game now)")
    except Exception as e:
        print("Failed to activate game window: {}".format(str(e)))
        print("Read the relevant restrictions here: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow")
        activationSuccess = False
        activationRetries = 0
        break
    # wait a little bit before the next try
    time.sleep(3.0)
    activationRetries = activationRetries - 1
# if we failed to activate the window then we'll be unable to send input to it
# so just exit the script now
if activationSuccess == False:
    exit()
print("Successfully activated the game window...")

# Attempt to open the serial port (uncomment when ready)
# arduino = serial.Serial('COM5', 9600, timeout=0)

left = ((videoGameWindow.left + videoGameWindow.right) // 2) - (ScreenSizeX // 2)
top = videoGameWindow.top + \
    (videoGameWindow.height - ScreenSizeY) // 2

with mss.mss() as sct:
    monitor = {'top': top, 'left': left, 'width': ScreenSizeX, 'height': ScreenSizeY}
    COLORS = np.random.uniform(0, 255, size=(1500, 3))
    while True:
        try:
            # Capture screen
            img = np.array(sct.grab(monitor))
        except mss.exception.ScreenShotError as e:
            print(f"ScreenShotError: {e}")
            break
        # Perform object detection
        results = model(img)
        rl = results.xyxy[0].tolist()

        # Process detection results
        if len(rl) > 0:
            if rl[0][4] >= 0.3:  # Confidence threshold
                if rl[0][5] == 1:  # Class prediction check
                    print(rl[0])
                    
                    # X-axis calculation
                    xmax = int(rl[0][2])
                    width = int(rl[0][2] - rl[0][0])
                    screenCenterX = ScreenSizeX / 2
                    centerX = int((xmax - (width / 2)) - screenCenterX)

                    # Y-axis calculation
                    ymax = int(rl[0][1])
                    height = int(rl[0][1] - rl[0][3])
                    screenCenterY = ScreenSizeY / 2
                    centerY = int((ymax - (height / 4)) - screenCenterY)

                    # Adjust movement
                    moveX = int(centerX * ampX * 0.1)
                    moveY = int(centerY * ampY * 0.1)

                    if centerY < screenCenterY:
                        moveY *= -1
                    
                    # Send movement commands to Arduino (uncomment when ready)
                    if win32api.GetKeyState(0x14):
                        # arduino.write((str(moveX) + ":" + str(moveY) + 'x').encode())
                        time.sleep(0.08)
        for i in range(0, len(rl)):
            halfW = round((rl[i][2] - rl[i][0]) / 2)
            halfH = round((rl[i][3] - rl[i][1]) / 2)
            midX = (rl[i][2] + rl[i][0]) / 2
            midY = (rl[i][3] + rl[i][1]) / 2
            (startX, startY, endX, endY) = int(
                midX + halfW), int(midY + halfH), int(midX - halfW), int(midY - halfH)

            idx = 0

            # draw the bounding box and label on the frame
            label = "{}: {:.2f}%".format(
                rl[i][6], rl[i][4] * 100)
            cv2.rectangle(img, (startX, startY), (endX, endY),
                          COLORS[idx], 2)
            y = startY - 15 if startY - 15 > 15 else startY + 15
            cv2.putText(img, label, (startX, y),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, COLORS[idx], 2)
        cv2.imshow('Live Feed', img)
        if (cv2.waitKey(1) & 0xFF) == ord('q'):
            exit()

        # Allow some time for the window to process the key press
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

cv2.destroyAllWindows()


#      xmin    ymin    xmax   ymax  confidence  class    name
# 0  749.50   43.50  1148.0  704.5    0.874023      0  person
# 1  433.50  433.50   517.5  714.5    0.687988     27     tie
# 2  114.75  195.75  1095.0  708.0    0.624512      0  person
# 3  986.00  304.00  1028.0  420.0    0.286865     27     tie
