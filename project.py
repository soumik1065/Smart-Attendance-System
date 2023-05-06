import face_recognition
import cv2
import numpy as np
from openpyxl import Workbook
import os
from datetime import datetime

video_capture = cv2.VideoCapture(0)

animesh_image = face_recognition.load_image_file("Dataset/Animesh.jpg")
animesh_encoding = face_recognition.face_encodings(animesh_image)[0]

soumik_image = face_recognition.load_image_file("Dataset/Soumik.jpg")
soumik_encoding = face_recognition.face_encodings(soumik_image)[0]

aritra_image = face_recognition.load_image_file("Dataset/Aritra.jpg")
aritra_encoding = face_recognition.face_encodings(aritra_image)[0]

arpan_image = face_recognition.load_image_file("Dataset/Arpan.jpg")
arpan_encoding = face_recognition.face_encodings(arpan_image)[0]

sayanik_image = face_recognition.load_image_file("Dataset/Sayanik.jpg")
sayanik_encoding = face_recognition.face_encodings(sayanik_image)[0]

subarna_image = face_recognition.load_image_file("Dataset/Subarna.jpg")
subarna_encoding = face_recognition.face_encodings(subarna_image)[0]

raja_image = face_recognition.load_image_file("Dataset/Raja.jpg")
raja_encoding = face_recognition.face_encodings(raja_image)[0]

swapnodip_image = face_recognition.load_image_file("Dataset/Swapnodip.jpg")
swapnodip_encoding = face_recognition.face_encodings(swapnodip_image)[0]

sahid_image = face_recognition.load_image_file("Dataset/Sahid.jpg")
sahid_encoding = face_recognition.face_encodings(sahid_image)[0]


know_face_encoding=[
animesh_encoding,
soumik_encoding,
aritra_encoding,
arpan_encoding,
sayanik_encoding,
subarna_encoding,
raja_encoding,
swapnodip_encoding,
sahid_encoding
]

know_faces_names=[
"Animesh",
"Soumik",
"Aritra",
"Arpan",
"Sayanik",
"Subarna",
"Raja",
"Swapnadip",
"Sahid"
]

students = know_faces_names.copy()

face_locations=[]
face_encodings=[]
face_names=[]
s= True

#excel file manipulation

wb = Workbook()

ws = wb.active
ws.title = "Data"


now=datetime.now()
current_date=now.strftime("%Y-%m-%d")


while True:
	_,frame = video_capture.read()
	small_frame = cv2.resize(frame,(0,0),fx=0.25,fy=0.25)
	rgb_small_frame = small_frame[:,:,::-1]
	if s:
		face_locations=face_recognition.face_locations(rgb_small_frame)
		face_encodings=face_recognition.face_encodings(rgb_small_frame,face_locations)
		face_names=[]
		for face_encoding in face_encodings:
			matches=face_recognition.compare_faces(know_face_encoding,face_encoding)
			name=""
			face_distance=face_recognition.face_distance(know_face_encoding,face_encoding)
			best_match_index=np.argmin(face_distance)
			if matches[best_match_index]:
				name= know_faces_names[best_match_index]

			face_names.append(name)
			if name in know_faces_names:
				if name in students:
					students.remove(name)
					print('Attendance marked!')
					current_time=now.strftime("%H-%M-%S")
					ws.append([name,current_time])
					wb.save(f'{current_date}.xlsx') 


	cv2.imshow("Attendence System",frame)
	if cv2.waitKey(1) & 0xFF == ord('q'):
		break

video_capture.release()
cv2.destroyAllWindows()
f.close( )



