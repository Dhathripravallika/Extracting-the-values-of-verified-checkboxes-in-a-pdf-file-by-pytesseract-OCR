## Extracting the values corresponding to the verified checkboxes in a pdf-file by pytesseract-OCR

In this repository, the code is written to extract the corresponding text values following the verified checkboxes in a given PDF file. The code first converts each page of the pdf file to the image, then detects the verified checkboxes and finally extracts the corresponding text beside the verified checkbox. Surprisingly the code is able to extract and show the parameter which was subjected to be verified, too. 
Here is an example of the output of the code I screenshoted and corresponding page in the PDF file.


![input](https://user-images.githubusercontent.com/128442592/236622776-e6c3292d-202a-47b5-9f01-63eacca61bc6.png)





![output](https://user-images.githubusercontent.com/128442592/236622976-bf534bd3-f614-474c-b573-7a7dbee6f2df.png)


Just to note that it might have some spelling errors which is due to the pretrained model used.
