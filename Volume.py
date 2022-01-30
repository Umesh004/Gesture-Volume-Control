import cv2
import time
import numpy as np
import math
import HandModule as hm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
wcam,hcam=640,480
cap=cv2.VideoCapture(0)
cap.set(3,wcam)
cap.set(4,hcam)
ptime=0
ctime=0
vol=0
volBar=400
volPer=0
detector=hm.handDetector(detectionCon=0.6)
# pycaw : dynamic link library in dll form so ctypes is used to cast it
# Accessing the speaker using AudioUtilities.GetSpeakers() method
devices = AudioUtilities.GetSpeakers()
# Activating speakers using Activate() method
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
#volume.GetMute()
#volume.GetMasterVolumeLevel()
print("Volume Range : ",volume.GetVolumeRange())
# Note min and max range of volume
volRange=volume.GetVolumeRange()
minvol=volRange[0]
maxvol=volRange[1]

while True:
    success,img=cap.read()
    img=detector.findHands(img)
    lmList=detector.findPosition(img,draw=False)
    if len(lmList)!=0:
        #print(lmList[4],lmList[8])
        x1,y1=lmList[4][1],lmList[4][2] #tip of thumb finger's x and y position
        x2,y2=lmList[8][1],lmList[8][2] # tip of index finger's x and y position
        cx,cy=(x1+x2)//2,(y1+y2)//2
        # highlighting required landmarks
        cv2.circle(img,(x1,y1),5,(255,0,255),cv2.FILLED)
        cv2.circle(img,(x2,y2),5,(255,0,255),cv2.FILLED)
        cv2.circle(img,(cx,cy),5,(255,0,255),cv2.FILLED)
        cv2.line(img,(x1,y1),(x2,y2),(255,0,255),3)
        # .hypot() gives mid point between 2 coordinates
        length=math.hypot(x2-x1,y2-y1)
        #print(length)
        #Hand Range 50 to 200
        #Volume Range -65 to 0
        #in order to match volume with the length interp
        vol=np.interp(length,[50,110],[minvol,maxvol]) # Conversion of volume in terms of length
        volBar=np.interp(length,[50,110],[400,150])
        volPer=np.interp(length,[50,110],[0,100])
        print(int(length),vol)
        volume.SetMasterVolumeLevel(vol, None) # using SetMasterVolumeLevel() set the device volume
        if length<35:
            cv2.circle(img,(cx,cy),7,(0,255,0),cv2.FILLED)
    cv2.rectangle(img,(50,150),(85,400),(0,255,0),3)
    cv2.rectangle(img,(50,int(volBar)),(85,400),(0,255,0),cv2.FILLED)
    cv2.putText(img,f'{int(volPer)} %',(50,450),cv2.FONT_HERSHEY_PLAIN,1,(255,0,0),2)
    ctime=time.time()
    fps=1/(ctime-ptime)
    ptime=ctime
    cv2.putText(img,f'FPS: {int(fps)}',(40,50),cv2.FONT_HERSHEY_PLAIN,1,(255,0,0),2)
    cv2.imshow("Image",img)
    k=cv2.waitKey(30) & 0xff
    if k==27:
          break
cap.release()
cv2.destroyAllWindows()
