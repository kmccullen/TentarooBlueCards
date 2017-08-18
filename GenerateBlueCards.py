#
# Tentaroo Blue Card merge/split/rename
#
# Take two input PDF files that contain front and back images of blue cards.
# Produces separate duplexed files containing each individual blue card.
# Individual files are named by Scout, Merit Badge, page number, and part of page (by thirds).
#
# This code is *highly* dependent upon three things:
# 1. The ability to extract the text from the PDF file
# 2. The exact arrangement of the text within that file (i.e. it searches for phrases that preceed the
#    Scout's name and the Merit Badge name.
# 3. Punctuation within a name and/or merit badge. You may find modifications necessary.
#
# Written By: Kevin McCullen
#             Green Mountain Council
#
# Copyright: 2017
#
# Licensed Under: The Unlicense and "The Scout Law"
#
# Disclaimer: No Support, No Warranty, No Recourse
#
# Requires: PyPDF2
#           xpdf-tools
#
import os
import subprocess
import shutil
from PyPDF2 import PdfFileWriter, PdfFileReader

workingDir = "C:/Users/Kevin/Documents/Scouts/2017BlueCards"
inFileFront = workingDir + "/" + "front.pdf"
inFileBack = workingDir + "/" + "back.pdf"
duplexedFile = workingDir + "/" + "duplexed.pdf"

event = "MNSR_2017_Week5"

pdf2TextPath = "c:/Users/Kevin/Downloads/xpdf-tools-win-4.00/xpdf-tools-win-4.00/bin64"
pdf2Text = pdf2TextPath + "/pdftotext.exe"

def MergeFrontBack():
    inFD1 = open(inFileFront, "rb")
    inPdf1 = PdfFileReader(inFD1)
    inFD2 = open(inFileBack, "rb")
    inPdf2 = PdfFileReader(inFD2)
    outPdf = PdfFileWriter()

    if (inPdf1.getNumPages() != inPdf2.getNumPages()):
        print("Front and back PDF files have different page counts!")
    else:
        for i in range(0, inPdf1.getNumPages()):
            front = inPdf1.getPage(i)
            outPdf.addPage(front)
            back = inPdf2.getPage(i)
            outPdf.addPage(back)

        outFD = open(duplexedFile, "wb")
        outPdf.write(outFD)
        outFD.close()

    inFD1.close()
    inFD2.close()

def CutIntoThirds():
    lowerThird = 275
    upperThird = 525
    top = 792
    yCoords = [0, lowerThird, upperThird, top]

    def ExtractOneThird(inPdf, j):
        print("Third: " + str(j) + " Pages: " + str(inPdf.getNumPages()))
        for i in range(0,inPdf.getNumPages(),2):

            outPdf = PdfFileWriter()
        
            front = inPdf.getPage(i)
            front.trimBox.lowerLeft = (0, yCoords[j])
            front.trimBox.upperRight = (612, yCoords[j+1])
            front.cropBox.lowerLeft = (0, yCoords[j])
            front.cropBox.upperRight = (612, yCoords[j+1])
            front.mediaBox.lowerLeft = (0, 0)
            front.mediaBox.upperRight = (612, top)
            outPdf.addPage(front)

            back = inPdf.getPage(i+1)
            back.trimBox.lowerLeft = (0, yCoords[j])
            back.trimBox.upperRight = (612, yCoords[j+1])
            back.cropBox.lowerLeft = (0, yCoords[j])
            back.cropBox.upperRight = (612, yCoords[j+1])
            back.mediaBox.lowerLeft = (0, 0)
            back.mediaBox.upperRight = (612, top)
            outPdf.addPage(back)

            outFD = open(workingDir + "/" + "out_" + str(i) + "_" + str(j) + ".pdf", "wb")
            outPdf.write(outFD)
            outFD.close()

    inFile = open(duplexedFile, "rb")
    inPdf = PdfFileReader(inFile)
    ExtractOneThird(inPdf, 0)
    inFile.close()
    
    inFile = open(duplexedFile, "rb")
    inPdf = PdfFileReader(inFile)
    ExtractOneThird(inPdf, 1)
    inFile.close()
    
    inFile = open(duplexedFile, "rb")
    inPdf = PdfFileReader(inFile)
    ExtractOneThird(inPdf, 2)
    inFile.close()

def RenameCards():
    files = os.listdir(workingDir)
    print("FileCount: " + str(len(files)))

    for fileName in files:
        if (fileName.find(".pdf") >= 0 and fileName.find("out_") >= 0):
            pageNumStart = fileName[4:]
            pageNum = pageNumStart[:fileName.find("_")-1]
            cropNum = fileName[len(fileName)-5]
            txtFile = fileName.replace(".pdf", ".txt")
            rc = os.system(pdf2Text + ' ' + workingDir + "/" + fileName + " " + workingDir + "/" + txtFile)
            txtFD = open(workingDir + "/" + txtFile)
            txt = txtFD.read()
            nameAt = txt.find("APPLICANT'S RECORD Name")
            typeAt = txt.find("has given me his completed application for the")
            typeEnds = txt.find("Merit Badge\n\nCompleted on")
            nameIs = txt[nameAt+24 : typeAt-2]
            typeIs = txt[typeAt+48 : typeEnds-1]
            nameIs = nameIs.replace(" ", "_")
            nameIs = nameIs.replace("\n", "")
            nameIs = nameIs.replace("\t", "")
            typeIs = typeIs.replace(".", "")
            typeIs = typeIs.replace(" ", "_")
            typeIs = typeIs.replace("\n", "")
            typeIs = typeIs.replace("\t", "")
            newName = event + "_" + nameIs + "_" + typeIs + "_" + str(pageNum) + "_" + str(cropNum) + ".pdf"
            print(fileName, newName)
            shutil.copy2(workingDir + "/" + fileName, workingDir + "/" + newName)
        
MergeFrontBack()
CutIntoThirds()
RenameCards()
