import pdfrw
import os
import PyPDF2
import sys
import time



dirPath = input('Enter directory path (e.g d:/pdf) of Pdf files: ')
#sys.stdout = open('output.txt', 'w')
k = 0
all_files = []
for dirpath, dirnames, filenames in os.walk(dirPath):
    for filename in [f for f in filenames if f.endswith(".pdf")]:
        all_files.append(os.path.join(dirpath, filename))

if len(all_files) == 0:
    print('No Pdf files available to remove comment. Try again.')
else:
    print('Files and their location available for Comment removal')
    print(all_files)
    print('')
    print('Processing Total ' + str(len(all_files)) + ' files. Wait for few seconds...')
    print('Waiting for Success.....')
    while k < len(all_files):
        input1 = PyPDF2.PdfFileReader(open(all_files[k], "rb"))
        nPages = input1.getNumPages()
        print('')
        print("File Name: " + all_files[k])
        print('Comments will appear below if available..')
        for i in range(nPages):
            page0 = input1.getPage(i)
            try:
                for annot in page0['/Annots']:
                    subtype = annot.getObject()['/Subtype']
                    if subtype == "/Text" or subtype == "/Highlight":
                        #time.sleep(1)
                        print(annot.getObject()['/Contents'])
                        #print(annot.getObject()['/Author'])
            except:
                # print('There is no Comment in this file: ' + all_files[k])
                pass
        reader = pdfrw.PdfReader(all_files[k])
        for p in reader.pages:
            if p.Annots:
                p.Annots = [a for a in p.Annots if a.Subtype == "/Link"]

        pdfrw.PdfWriter(all_files[k], trailer=reader).write()
        k += 1
    print('')
    # sys.stdout.close()
    print('HURRAY! Comments extracted and removed Successfully (refer output.txt). See you later!')


