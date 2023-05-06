import numpy as np  # import NumPy library for numerical operations
import pytesseract  # import PyTesseract library for optical character recognition (OCR)
import pandas as pd  # import Pandas library for data manipulation and analysis
import cv2  # import OpenCV library for image processing
from pdf2image import convert_from_path  # import pdf2image library to convert pdf pages to images
import os  # import os module for file and directory manipulation
from PIL import Image  # import Image class from PIL library for image handling
from pytesseract import image_to_string  # import image_to_string function from PyTesseract library for OCR

from module import extract_boxes, extract_checkboxes  # import extract_boxes and extract_checkboxes functions from a module

images = convert_from_path('C:/Users/Taraneh/Desktop/pythonProject/ocr/pages/typeA1.pdf',  # read and convert pdf file to images
                           poppler_path='C:/poppler-0.68.0/bin')  # specify the path of the poppler library

path = "pages"  # create a variable to store the path of the directory to save images

for i in range(len(images)):  # loop over each image
    name = f'page' + str(i) + '.jpg'  # create a name for each image
    filename = os.path.join(path, name)  # create the path and name for each image
    images[i].save(filename)  # save the image to the specified path and name

a,b= [],[]  


for root, dirs, files in os.walk(path):  # loop over the path of the directory containing the images
    for filename in files:  # loop over each file in the directory
        final = {}  # create an empty dictionary to store the results
        text_list = []  # create an empty list to store text extracted from the image
        pg_nmb = filename.split(".")  # split the filename to extract the page number

        if pg_nmb[1] == 'jpg':  # check if the file is an image
            pg_nmb = pg_nmb[0]  # set the page number

            file_path = os.path.join(root, filename)  # create the path and name for the image file
            checkboxes = extract_checkboxes(file_path)  # extract checkboxes coordinates from the image
            boxes = extract_boxes(file_path)  # extract table cell coordinates from the image
            results = pytesseract.image_to_data(file_path, output_type=pytesseract.Output.DICT,  # extract text and coordinates from the image using OCR
                                                config='--psm 12 --oem 3 ')
            results_df = pd.DataFrame(results)  # convert the extracted data to a Pandas DataFrame
            for k, checkbox in enumerate(checkboxes):  # loop over each checkbox
                if checkbox[1] == 1:  # check if the checkbox is selected

                    text_list = []  # create an empty list to store text extracted from the corresponding table cell
                    b = []  # create an empty tuple to store checkbox and corresponding table cell
                    for j in range(0, len(results_df)):  # loop over each row in the DataFrame
                        x, y, w, h = results_df["left"][j], results_df["top"][j], results_df["width"][j], \
                            results_df["height"][j]  # get the location and size of the OCR text

                        if (k + 1) != len(checkboxes): #Check if the current checkbox is not the last checkbox on the page.
                            next_box = checkboxes[k + 1] # Store the coordinates of the next checkbox on the page.
                            if checkbox[0][1] - 10 <= next_box[0][1] <= checkbox[0][1] + 10: #Check if the y-coordinate of the next checkbox is within 10 pixels of the y-coordinate of the current checkbox.
                                if (checkbox[0][1] <= y <= checkbox[0][1] + checkbox[0][2]) and (
                                        checkbox[0][0] + checkbox[0][2] <= x <= next_box[0][0]): #Check if the current checkbox's y-coordinate is within its bounding box's height and if the x-coordinate is within the bounding boxes of the current and next checkbox. If it satisfies the condition, append the 
                                    text_list.append(results_df['text'][j]) #corresponding text value to the text_list.
                                    # print(text_list)
                            else:   #If the current checkbox does not meet the above condition, check if the current checkbox's y-coordinate is within its bounding box's height and the x-coordinate is within the current checkbox's bounding box. If it satisfies the condition, append the corresponding text value to the text_list.
                                if (checkbox[0][1] <= y <= checkbox[0][1] + checkbox[0][2]) and (
                                        checkbox[0][0] + checkbox[0][2] <= x):
                                    text_list.append(results_df['text'][j])  
                                  

                        if (k + 1) == len(checkboxes) and (checkbox[0][1] <= y <= checkbox[0][1] + checkbox[0][2]) and ( #If the current checkbox is the last checkbox on the page and its y-coordinate is within its bounding box's height and the x-coordinate is within the checkbox's bounding box, append the corresponding text value to the text_list.
                                checkbox[0][0] + checkbox[0][2] <= x):
                            text_list.append(results_df['text'][j])
                            # print(text_list)
                    text_list = [" ".join(text_list)]
               
                    a = (k, text_list) #Store the index of the current checkbox and the corresponding text value as a tuple in a.
                  
                    for ff, box in enumerate(boxes): # Loop through all boxes (including checkboxes) on the page.
                        if (box[0] < checkbox[0][0] < box[0] + box[2]) and (  #If a box contains the current checkbox, and the next box on the same row has the same y-coordinate, extract text from the region between the two boxes using OpenCV and Tesseract OCR.
                                box[1] < checkbox[0][1] < box[1] + box[3]) and \
                                boxes[ff + 1][1] == box[1]:
                            param = boxes[ff + 1]
                            im = cv2.imread(file_path)
                            imm = im[param[1]:param[1] + param[3], param[0]:param[0] + param[2]]
                            im2 = Image.fromarray(imm.astype('uint8'))
                            st = image_to_string(im2)
                            b = (k, st) #Store the index of the current checkbox and the extracted text value as a tuple in b.
                           

                    if b!=[]:  # Check if any text was extracted for the current checkbox.
                        print(f"\033[1m{b[1]}:{text_list}\033[0m")
                    else:  If no text was extracted for the current checkbox, print only the corresponding checkbox text enclosed in bold tags.
                        print(f"\033[1m{text_list}\033[0m")

                    print("-------------------------------------------------------------")  # Print a separator after each checkbox's text has been printed.
