import cv2
import mediapipe as mp
import time
class handDetector():
    # For .Hands() function initializing the parameters
    def __init__(self, mode=False, maxHands=1, detectionCon=0.5, trackCon=0.5):
        self.mode = mode
        self.maxHands = maxHands
        self.detectionCon = detectionCon
        self.trackCon = trackCon
        self.mpHands = mp.solutions.hands
        self.hands = self.mpHands.Hands(self.mode, self.maxHands,
                                        self.detectionCon, self.trackCon)
        self.mpDraw = mp.solutions.drawing_utils

    def findHands(self, img, draw=True):
        # BGR must be converted to RGB to work with media-pipe
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        self.results = self.hands.process(imgRGB)
        # print x,y,z coordinates of a unknown landmark
        #print(self.results.multi_hand_landmarks)

        if self.results.multi_hand_landmarks:
            for handLms in self.results.multi_hand_landmarks:
                # For loop is required for detection of each single hand.
                if draw:
                    # This will draw the landmarks and give connections on the hand.
                    self.mpDraw.draw_landmarks(img, handLms,
                                               self.mpHands.HAND_CONNECTIONS)
                    #mpHands.HAND_CONNECTIONS will give connections between the landmarks
        return img
    def findPosition(self, img, handNo=0, draw=True):

        lmList = []
        if self.results.multi_hand_landmarks:
            myHand = self.results.multi_hand_landmarks[handNo]
            for id, lm in enumerate(myHand.landmark):
                # print(id, lm) // Prints id and x,y,z position of respective landmarks
                #conversion decimal to pixel
                h, w, c = img.shape
                cx, cy = int(lm.x * w), int(lm.y * h)  # multiply coordinate position by height and width
                # print(id, cx, cy)
                lmList.append([id, cx, cy])
                if draw:
                    cv2.circle(img, (cx, cy), 6, (255, 0, 255), cv2.FILLED)

        return lmList
def main():
    pTime = 0
    cTime = 0
    cap = cv2.VideoCapture(0)
    detector = handDetector()
    while True:
        success, img = cap.read()
        # Process the hands first
        img = detector.findHands(img)
        # Draw landmarks and connect them
        lmList = detector.findPosition(img)
        #print(lmList)
        if len(lmList) != 0:
           print(lmList[4])
       # Calucating frame per second
        cTime = time.time()
        fps = 1 / (cTime - pTime)
        pTime = cTime
       # Placing calucated fps into VideoCapture() interface
        cv2.putText(img, str(int(fps)), (10, 70), cv2.FONT_HERSHEY_PLAIN, 3,
                    (255, 0, 255), 3)
       # Displaying the Video
        cv2.imshow("Image", img)
        k=cv2.waitKey(30) & 0xff
        if k==27:
          break
    cap.release()
cv2.destroyAllWindows()
if __name__ == "__main__":
    main()