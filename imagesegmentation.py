import argparse
import cv2
import os
import random


def imagesegmentation(img, path):
    r=random.randint(1000,60000)
    img = cv2.imread(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.threshold(blurred, 60, 255, cv2.THRESH_BINARY)[1]

    ksize = 5
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (ksize, ksize))
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

    cnts = cv2.findContours(thresh.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    # ~ cnts = cnts[0] if imutils.is_cv2() else cnts[1]
    cnts = cnts[1]
    minArea = 500
    count = 0
    filename = ""
    for c in cnts:
        area = cv2.contourArea(c)
        if area > minArea:
            c = c.astype("float")
            c = c.astype("int")
            # cv2.drawContours(img, [c], -1, (0, 255, 0), 2)

            x, y, w, h = cv2.boundingRect(c)
            if w < 100 or h < 100:
                continue  # continue here

            count = count + 1
            cv2.imwrite(os.path.join(path, ('tmpimg_'+str(r)+ str(count) + '.jpg')), img[y:y + h, x:x + w])

            if len(filename) == 0:
                filename = filename + ('tmpimg_'+str(r)+ str(count) + '.jpg')
            else:
                filename = filename + "|" + ('tmpimg_'+str(r)+ str(count) + '.jpg')

    if len(filename) == 0:
        cv2.imwrite(os.path.join(path, ('tmpimg_'+str(r)+ str(1) + '.jpg')), img[:, :])
        filename = ('tmpimg_'+str(r)+ str(count) + '.jpg')

    print(filename)

    # cv2.imshow("Image", img)
    cv2.waitKey(0)
    return filename


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--image", required=True, help="Path to the image")
    ap.add_argument("-p", "--paths", required=True, help="Path to store image")
    args = vars(ap.parse_args())

    imagesegmentation(args["image"], args["paths"])


if __name__ == '__main__':
    main()
