import sys
import cv2
import pytesseract
import PySimpleGUI as sg
import os.path
import docx 
import re
from tableDetect import tableDetectExtract

#if TesseractNotFoundError occurs
#pytesseract.pytesseract.tesseract_cmd = r'path\to\tesseract.exe'

def imgTextConvert(image, folderChosen, fileTypeChoice):
    image = cv2.imread(image)
    image = cv2.resize(image, None, fx=1.2, fy=1.2, interpolation=cv2.INTER_CUBIC)
    #image = cv2.blur(image,(1,1))
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    image = cv2.threshold(image, 0,255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    #image = cv2.medianBlur(image, 3)

    textConverted = pytesseract.image_to_string(image)
    textConverted = re.sub(r'[^\x00-\x7F]+|\x0c',' ', textConverted)

    if fileTypeChoice == 0:
        f = open(folderChosen + "/Exported Text.txt", "w")
        f.write( textConverted)
        f.close()
    else:
        exportFolder = os.path.join(folderChosen, 'Exported Text.docx')
        document.add_paragraph(textConverted)
        document.save(exportFolder)

def imgDetectionCrop(imgChosen, folderChosen, fileTypeChoice):
    if fileTypeChoice == 0:
        fileType = '.jpg'       
    else:
        fileType = '.png'
    img = cv2.imread(imgChosen) 
    gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY) 
    edged = cv2.Canny(img, 10, 250) 
    (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE) 
    idx = 0 
    for c in cnts: 
        x,y,w,h = cv2.boundingRect(c) 
        
        if w>100 and h>100:
            idx +=1 
            new_img=img[y:y+h,x:x+w]
            if fileTypeChoice == 0: 
                cv2.imwrite(folderChosen + '/JPG Export ' + str(idx) + fileType, new_img)
            else:
                cv2.imwrite(folderChosen + '/PNG Export ' + str(idx) + fileType, new_img)

def main(title):
    layout = [  [sg.Text('Please choose an image you wish to process :' ,size=(34,None)), sg.Input(), sg.FileBrowse(file_types=(("*.png *.jpg *.jpeg *.webp", "*.png *.jpg *.jpeg *.webp"),), key = 'imgChosen')],
                [sg.Text('Please choose the export folder :', size=(34,None)), sg.Input(), sg.FolderBrowse(key = 'folderChosen')],
                [sg.Button('Submit')],
                [sg.Text(' ')],
                [sg.Text('Save image(s) as :' ,size=(14,None) ), sg.Checkbox('Do not save', key='inputNoSave1', size=(8,None)), sg.Checkbox('*jpg', key='inputJpgSave', size=(3,None)), sg.Checkbox('*.png', key='inputPngSave')],
                [sg.Text('Save text as :',size=(14,None)), sg.Checkbox('Do not save',  key='inputNoSave2', size=(8,None)), sg.Checkbox('*txt',  key='inputTxtSave', size=(3,None)), sg.Checkbox('*.docx',  key='inputDocxSave')],
                [sg.Text('Save table(s) as :',size=(14,None)), sg.Checkbox('Do not save',  key='inputNoSave3', size=(8,None)), sg.Checkbox('*csv',  key='inputCsvSave', size=(3,None)), sg.Checkbox('*.xlsx',  key='inputXlsxSave')],
                [sg.Button('Run'), sg.Text('   '),sg.Button('Exit')]
    ]
    return sg.Window(title, layout)

document = docx.Document()
submitted = 0

#Calls window
mainWindow = main(' ')

while True:
    
    while True:
        event, values = mainWindow.read()
        #user presses 'Exit' or closes the window
        if event == sg.WINDOW_CLOSED or event== 'Exit':
            sys.exit()

        #user presses 'Submit'
        elif event == 'Submit':
            imgChosen = values['imgChosen']
            folderChosen = values['folderChosen']
            submitted = 1
            sg.popup("Image & Folder Location Submitted")
        
        #Checks each radio option once user presses run
        elif event == 'Run' and submitted == 1:

            if values['inputNoSave1']== True and values['inputNoSave2']== True and values['inputNoSave3']== True:
                sg.popup('Invalid choices')

            if values['inputNoSave1']== False:
                if values['inputJpgSave'] == True:
                    imgDetectionCrop(imgChosen, folderChosen, 0)
                
                if values['inputPngSave'] == True:
                    imgDetectionCrop(imgChosen, folderChosen, 1)

            if values['inputNoSave2']== False:
                if values['inputTxtSave'] == True:
                    imgTextConvert(imgChosen, folderChosen, 0)
                
                if values['inputDocxSave'] == True:
                    imgTextConvert(imgChosen, folderChosen, 1)

            if values['inputNoSave3']== False:
                if values['inputCsvSave'] == True:
                    tableDetectExtract(imgChosen, folderChosen, 0)

                if values['inputXlsxSave'] == True:
                    tableDetectExtract(imgChosen, folderChosen, 1)

            submitted = 0
            sg.popup('Extraction Complete\nFiles have been saved at', folderChosen)    
            break
        
        #User presses 'Run' but haven't submit their images
        elif event == 'Run' and submitted==0:
            sg.popup("Please press the 'Submit' button to continue")
