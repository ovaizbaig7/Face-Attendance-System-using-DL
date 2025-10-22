from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video=cv2.VideoCapture(0)
facedetect=cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS=pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES=pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

knn=KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

COL_NAMES = ['DATE', 'NAME', 'TIME']

attendance = []
attended_today = set()

while True:
    ret,frame=video.read()
    gray=cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces=facedetect.detectMultiScale(gray, 1.3,5)
    for (x,y,w,h) in faces:
        crop_img=frame[y:y+h, x:x+w, :]
        resized_img=cv2.resize(crop_img, (50,50)).flatten().reshape(1,-1)
        output=knn.predict(resized_img)
        ts=time.time()
        date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp=datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        if str(output[0]) not in attended_today:
            cv2.rectangle(frame, (x,y), (x+w, y+h), (0,0,255), 1)
            cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),2)
            cv2.rectangle(frame,(x,y-40),(x+w,y),(50,50,255),-1)
            cv2.putText(frame, str(output[0]), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
            attendance.append([date, str(output[0]), timestamp])
            attended_today.add(str(output[0]))
            cv2.rectangle(frame, (x,y), (x+w, y+h), (0,0,255), 1)
            speak("Camera is on for attendance..")
            time.sleep(2)
    cv2.imshow("Attendance", frame)
    k=cv2.waitKey(1)
    if k==ord('o'):
        speak("Attendance Taken..")
        time.sleep(1)
    
        if os.path.isfile("Attendance/Attendance_" + date + ".csv"):
            exists = True
        else:
            exists = False

        if exists:
            with open("Attendance/Attendance_" + date + ".csv", "a", newline='') as csvfile:
                writer=csv.writer(csvfile, delimiter=',')
                writer.writerows(attendance)
        else:
            with open("Attendance/Attendance_" + date + ".csv", "w", newline='') as csvfile:
                writer=csv.writer(csvfile, delimiter=',')
                writer.writerow(COL_NAMES)
                writer.writerows(attendance)
        attendance = []
        break
    if k==ord('q'):
        break

video.release()
cv2.destroy