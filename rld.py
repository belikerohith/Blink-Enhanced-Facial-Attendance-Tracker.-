from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime

from win32com.client import Dispatch
import dlib
from scipy.spatial import distance as dist

def shape_to_landmarks(shape):
    landmarks = []
    for i in range(0, 68):
        x = shape.part(i).x
        y = shape.part(i).y
        landmarks.append((x, y))
    return landmarks

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

def eye_aspect_ratio(eye):
    A = dist.euclidean(eye[1], eye[5])
    B = dist.euclidean(eye[2], eye[4])
    C = dist.euclidean(eye[0], eye[3])
    ear = (A + B) / (2.0 * C)
    return ear

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)



COL_NAMES = ['NAME', 'TIME']

# Blink detection constants and parameters
EYE_AR_THRESH = 0.3
EYE_AR_CONSEC_FRAMES = 2
COUNTER = 0
TOTAL = 0
attendance_taken = False

detector = dlib.get_frontal_face_detector()
predictor = dlib.shape_predictor('data/shape_predictor_68_face_landmarks.dat')

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    
    # Detect faces and perform blink detection
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        
        rects = detector(gray, 0)
        for rect in rects:
            shape = predictor(gray, rect)
            
            landmarks = shape_to_landmarks(shape)
            left_eye = landmarks[36:42]
            right_eye = landmarks[42:48]
            ear_left = eye_aspect_ratio(left_eye)
            ear_right = eye_aspect_ratio(right_eye)
            
            if not attendance_taken:
                if ear_left < EYE_AR_THRESH or ear_right < EYE_AR_THRESH:
                    COUNTER += 1
                else:
                    if COUNTER >= EYE_AR_CONSEC_FRAMES:
                        TOTAL += 1
                        print("Blink detected")
                        # Take attendance or perform the desired action here
                        attendance_taken = True
                    COUNTER = 0
                
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (0, 255, 0), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 1)
        attendance = [str(output[0]), str(timestamp)]
        
    
    cv2.imshow("Frame", frame)
    
    k = cv2.waitKey(1)
    if k == ord('o') and TOTAL > 0:
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()
        attendance_taken = False
        COUNTER = 0
        TOTAL = 0
    elif k == ord('o') and TOTAL == 0:
        speak("No blink detected. Attendance not taken.")
    if k == ord('q'):
        break
        
video.release()
cv2.destroyAllWindows()
