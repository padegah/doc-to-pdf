from docx import Document
import os
import docx
import comtypes.client
from docx2pdf import convert
import re
import shutil
import docx2pdf
import time

# Get the list of all files and directories
path = "C://Users//PrinceKofiAdegah//receipt//files//"
output = "C://Users//PrinceKofiAdegah//receipt//output//"
dir_list = os.listdir(path)

def getFilename(file_path):
    new_file_name = ""
    with open (file_path) as f:

        pathname, extension = os.path.splitext(file_path)
        filename = pathname.split('/')

        if (extension == ".txt"):
            first_line = f.readline()
            line = first_line.split()
            new_file_name = filename[-1]+"-"+line[-1]
    
    return new_file_name


def getFileContent(file_path):
    content = ""
    updatedContent = ""
    tempContent = []
    tempDetails = []

    pathname, extension = os.path.splitext(file_path)
    with open (file_path) as f:

        if (extension == ".txt"):
            content = f.readlines()
            # print(content)
            for line in content:
                if "Desktop Test" in line:
                    line = re.sub("Desktop Test", "Systems", line)
                if "CASH" in line:
                    line = re.sub("CASH", "BANK SWIFT", line)
                if "Amount" in line:
                    tempDetails.append(line)
                if "Meter" in line:
                    tempDetails.append(line)
                if "Account" in line:
                    tempDetails.append(line)
                if "Customer" in line:
                    tempDetails.append(line)
                tempContent.append(line)
    updatedContent = ''.join(tempContent)
    
    return updatedContent, tempDetails


def updateTemplate(content, filename, base_path, output_pat):
    template = "template.docx"
    document  = Document(base_path+template)

    Table = document.tables[0]

    for cells in Table.rows[-2].cells:
        cells.text = content
        break

    document.save(base_path+filename+".docx")

    shutil.move(base_path+filename+".docx", output_pat+filename+".docx")


for ffile in dir_list:
    file_path = path+ffile
    filename = getFilename(file_path)
    content, tDetails = getFileContent(file_path)
    updateTemplate(content, filename, path, output)
    print(filename)
    # print(output+filename+".docx")

    data = ''.join(tDetails)

    file = open('items.txt','a')
    file.write(data)
    file.write("::::::\n")
    file.close
    
    docx2pdf.convert(output+filename+".docx")