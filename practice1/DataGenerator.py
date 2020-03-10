import os
import cv2
import time
import sys
import face_recognition
import glob
vid_cam = cv2.VideoCapture(0)
cascPath = 'haarcascade_frontalface_alt.xml'
faceCascade = cv2.CascadeClassifier(cascPath)
process_this_frame = True


def newLabelFolder(name):
    #cd = current directory
    #fd = final directory
    cd = os.getcwd()
    fd = os.path.join(cd, 'known_people', name)
    if not os.path.exists(fd):
        os.makedirs(fd)
    return fd

def detection(process_this_frame):
    name = ""
    while True:
        # Grab a single frame of video
        ret, frame = vid_cam.read()

        # Resize frame of video to 1/4 size for faster face recognition processing
        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

        # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
        rgb_small_frame = small_frame[:, :, ::-1]

        # Only process every other frame of video to save time
        if process_this_frame:
            # Find all the faces and face encodings in the current frame of video
            face_locations = face_recognition.face_locations(rgb_small_frame)
        process_this_frame = not process_this_frame
        # Display the results
        for (top, right, bottom, left) in face_locations:
            # Scale back up face locations since the frame we detected in was scaled to 1/4 size
            top *= 4
            right *= 4
            bottom *= 4
            left *= 4
            # Draw a box around the face
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
            # Display the resulting image
            cv2.imshow('Video', frame)
            k = cv2.waitKey(1) & 0xFF
            # Hit 'q' on the keyboard to quit!
            if k == 27 or k == 13:
                quit()
            elif k == ord('c') or k == ord('C'):
            #To save cropped face in a new/existing label
                print('Enter the full name of this user: ', end="")
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                newlabel = input()
                print('Enter the Number of this user: ', end="")
                number=input()
                label=newlabel.split(" ")[0].lower()
                wd = newLabelFolder(label)
                files = os.listdir(wd)
                os.chdir(wd)
                jpgCounter = len(glob.glob1(os.getcwd(), "*.jpg"))
                print (jpgCounter)
                if jpgCounter:
                    path = os.path.join(wd, label+str(jpgCounter)+".jpg")
                else:
                    path = os.path.join(wd, label+".jpg")
                text_file = os.path.join(wd, label+".txt")

                file_object = open(text_file, "w")
                file_object.write(newlabel+"\n") 
                file_object.writelines(number)
                file_object.close()

                cv2.imwrite(path, gray[top:bottom, left:right])
    return name

detection(process_this_frame)
