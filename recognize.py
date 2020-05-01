#Step 3
#To Execute
#python3 --encodings   __________.pickle  --image _______


import pandas as pd 
import os

import face_recognition
import argparse
import pickle
import cv2


ap = argparse.ArgumentParser()
ap.add_argument("-e", "--encodings", required=True,
	help="path to serialized db of facial encodings")
ap.add_argument("-i", "--image", required=True,
	help="path to input image")
ap.add_argument("-d", "--detection-method", type=str, default="hog",
	help="face detection model to use: either `hog` or `cnn`")
args = vars(ap.parse_args())


print("[INFO] loading encodings...")
data = pickle.loads(open(args["encodings"], "rb").read())


image = cv2.imread(args["image"])
rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)


print("[INFO] recognizing faces...")
boxes = face_recognition.face_locations(rgb,
	model=args["detection_method"])
encodings = face_recognition.face_encodings(rgb, boxes)


names = []


for encoding in encodings:

	matches = face_recognition.compare_faces(data["encodings"],
		encoding)
	name = "Unknown"

	
	if True in matches:
		matchedIdxs = [i for (i, b) in enumerate(matches) if b]
		counts = {}

		for i in matchedIdxs:
			name = data["names"][i]
			counts[name] = counts.get(name, 0) + 1

		name = max(counts, key=counts.get)
	
	names.append(name)


for ((top, right, bottom, left), name) in zip(boxes, names):

	cv2.rectangle(image, (left, top), (right, bottom), (0, 255, 0), 2)
	y = top - 15 if top - 15 > 15 else top + 15
	cv2.putText(image, name, (left, y), cv2.FONT_HERSHEY_SIMPLEX,
		0.75, (0, 255, 0), 2)


cv2.imshow("Image", image)
cv2.waitKey(0)

#################################################################################
                        #CREATING EXCEL FILE
##############################################################################
dataframe=pd.DataFrame(  {"roll_no" : ["218" ,"234 ","240"] ,"name":["ashu" , "tejas","ninad"],   "Attend" : 0 }  )

if os.path.exists("attendance.xlsx") ==  False :

    writer = pd.ExcelWriter('attendance.xlsx' , engine='xlsxwriter')

    dataframe.to_excel(writer, sheet_name='Sheet1', index=None)   
    print(dataframe)
    writer.save()

else:
    data=pd.read_excel('attendance.xlsx' , index_col=None)   
    dic={}
    for inx in data.index:
        dic[data.iloc[inx][1]]=data.iloc[inx][2]


    for key , value in dic.items():
        if key in names:
            dic[key]= value+1
                
    for inx ,val  in enumerate(dic.values()):
          data.iat[inx , 2]=val  
          
    #print(data)    
    data.to_excel("attendance.xlsx", sheet_name='Sheet1',index=False)
    print(data)
