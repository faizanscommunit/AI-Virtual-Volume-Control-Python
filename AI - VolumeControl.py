# Imports
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import math
import cv2
import time
import numpy as np
from cvzone.HandTrackingModule import HandDetector

# Variables
wCam, hCam = 640, 480

cap = cv2.VideoCapture(1)
cap.set(3, wCam)
cap.set(4, hCam)

detector = HandDetector(detectionCon=0.5)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))


volumeRange = volume.GetVolumeRange()
minVol = volumeRange[0]
maxVol = volumeRange[1]
vol = 0



while True:
    success, img = cap.read()
    hands = detector.findHands(img)
    lmList = detector.findPosition(hands, draw=False)[0]
    if len(lmList)!=0:
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]

        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
        cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)


        length = math.hypot(x2-x1, y2-y1)
        # print(length)

        vol = np.interp(length, [15, 150], [minVol, maxVol])
        volBar = np.interp(length, [15, 150], [400, 150])
        volText = np.interp(length, [15, 150], [0, 100])
        volume.SetMasterVolumeLevel(vol, None)

        cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
        cv2.rectangle(img, (50, int(volBar)), (85, 400), (0, 255, 0), cv2.FILLED)

        cv2.putText(img, f'Vol: {int(volText)} %', (30, 50),
                    cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)
        if length<50:
            cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)


    cv2.imshow("AI - Volume Controller", img)
    if cv2.waitKey(10) == ord('q'):
        break
