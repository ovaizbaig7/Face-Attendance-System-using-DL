import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
import threading

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

def process_frame(frame, knn, facedetect, LABELS, attended_today, attendance):
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
            cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),thickness=2)
            cv2.putText(frame, str(output[0]), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
            attended_today.add(str(output[0]))
            attendance.append([str(output[0]), str(date), str(timestamp)])

        elif str(output[0]) in attended_today:
            cv2.rectangle(frame, (x,y), (x+w, y+h), (0,255,0), 1)
            cv2.putText(frame, str(output[0]), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
            speak("Camera is on for attendance..")
            time.sleep(2)

video=cv2.VideoCapture(0)
facedetect=cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS=pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    face_recognizer=pickle.load(f)

attend_data=[]
attend_today=set()
attendance=[]

while True:
    ret, frame=video.read()
    if ret:
        threading.Thread(target=process_frame, args=(frame, face_recognizer, facedetect, LABELS, attend_today, attendance)).start()
        cv2.imshow('Recognizing...', frame)
        if cv2.waitKey(1) & 0xFF==ord('q'):
            break

video.release()
cv2.destroyAllWindows()

for data in attendance:
    attend_data.append(data)

with open('data/attendance_data.csv', 'w', newline='') as f:
    writer=csv.writer(f)
    writer.writerows(attend_data)

speak("The Attendance Sheet is Updated")
time.sleep(2)
speak("Exiting...")
time.sleep(2)