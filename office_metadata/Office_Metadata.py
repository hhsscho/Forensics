import os
import docx
import csv
import hashlib
from openpyxl import load_workbook
from datetime import datetime
from docx import Document
from pptx import Presentation

#"""""""""""""""Calculate File Size"""""""""""""""
def CalculateFileSize(file_size):
    if file_size >= (1024 * 1024):
        file_size = file_size / 1024 / 1024
    else:
        file_size = file_size / 1024
    return file_size

#"""""""""""""""Calculate Disk Allocation Size"""""""""""""""
def AllocationSize(allocation_size):
    if allocation_size % 1024 == 0:
        if allocation_size >= (1024 * 1024):
            allocation_size = allocation_size / 1024 / 1024
        else:
            allocation_size = allocation_size / 1024
    else:
        remainder = allocation_size % 1024
        allocation_size = allocation_size + 1024 - remainder
        if (allocation_size / 1024) % 4 == 0:
            if allocation_size >= (1024 * 1024):
                allocation_size = allocation_size / 1024 / 1024
            else:
                allocation_size = allocation_size / 1024
        else:
            allocation_size = allocation_size + 1024 * (4 - (allocation_size / 1024) % 4)
            if allocation_size >= (1024 * 1024):
                allocation_size = allocation_size / 1024 / 1024
            else:
                allocation_size = allocation_size / 1024
    return allocation_size

Path = os.getcwd()
files = os.listdir() #All File List
print("Current Path : ", Path, '\n')

#"""""""""""""""""""""XLSX FILE METADATA PARSER"""""""""""""""""""""
#Make 'XLSX_Metadata.csv' File
with open('XLSX_Metadata.csv', 'w', encoding='euc-kr', newline = '') as x_file:
    writer = csv.writer(x_file)
    writer.writerow(['Number', "FileName", "FileType", "Created Date", "Created Time(UTC +0:00)",
                "Modified Date", "Modified Time(UTC +0:00)", "File Size", "Disk Allocation Size",
                "Creator", "Last Modified By", "Version", "MD-5", "SHA-1"])

# Choose XLSX Files
xlsx_file = []
for x in files:
    if x.endswith('xlsx'):
        xlsx_file.append(x)
print('XlSX Found:', len(xlsx_file), 'Files')
print('XLSX File Metadata Exporting...')
for x in range(0, len(xlsx_file)):
    #Collect XLSX File Properties
    prop = load_workbook(xlsx_file[x]).properties
    Created_time = os.path.getctime(xlsx_file[x])
    Modified_time = os.path.getmtime(xlsx_file[x])

    f = open(xlsx_file[x], 'rb')
    data = f.read()
    f.close()

    #Get File Size
    file_size = os.path.getsize(xlsx_file[x])
    if file_size >= (1024 * 1024):
        Unit = "MB"
    else:
        Unit = "KB"
    x_bytesize = file_size
    allocation_size = file_size

    #Calculate File Size
    Alled_file_size = str(AllocationSize(allocation_size))[:5]
    #Calculate Disk Alloction size
    Caled_file_size = str(CalculateFileSize(file_size))[:5]

    row = [x+1, xlsx_file[x],os.path.splitext(xlsx_file[x])[-1],
                datetime.utcfromtimestamp(Created_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Created_time).strftime('%H:%M:%S'),
                datetime.utcfromtimestamp(Modified_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Modified_time).strftime('%H:%M:%S'),
                Caled_file_size + Unit, Alled_file_size + Unit,
                prop.creator, prop.lastModifiedBy, prop.version,
                hashlib.md5(data).hexdigest(),
                hashlib.sha1(data).hexdigest()]

    with open('XLSX_Metadata.csv', 'a', newline = '') as x_file:
        writer = csv.writer(x_file)
        writer.writerow(row)
    x_file.close()
print('XLSX File Metadata Exporting Finished', '\n')

#"""""""""""""""""""""DOCX FILE METADATA PARSER"""""""""""""""""""""
#Make 'DOCX_Metadata.csv' File
with open('DOCX_Metadata.csv', 'w', encoding='euc-kr', newline = '') as d_file:
    writer = csv.writer(d_file)
    writer.writerow(['Number', "FileName", "FileType", "Created Date", "Created Time(UTC +0:00)",
                "Modified Date", "Modified Time(UTC +0:00)", "File Size", "Disk Allocation Size",
                "Creator", "Last Modified By", "Version", "MD-5", "SHA-1"])

#Choose DOCX Files
docx_file = []
for x in files:
    if x.endswith('docx'):
        docx_file.append(x)
print('DOCX Found:', len(docx_file), 'Files')
print('DOCX File Metadata Exporting...')
for x in range(0, len(docx_file)):
    #Collect DOCX File Properties
    Created_time = os.path.getctime(docx_file[x])
    Modified_time = os.path.getmtime(docx_file[x])
    d_file = Document(str(docx_file[x]))
    properties = d_file.core_properties

    f = open(docx_file[x], 'rb')
    data = f.read()
    f.close()

    #Get File Size
    file_size = os.path.getsize(docx_file[x])
    if file_size >= (1024 * 1024):
        Unit = "MB"
    else:
        Unit = "KB"
    d_bytesize = file_size
    allocation_size = file_size

    #Calculate File Size
    Alled_file_size = str(AllocationSize(allocation_size))[:5]
    #Calculate Disk Alloction size
    Caled_file_size = str(CalculateFileSize(file_size))[:5]

    if properties.version == '':
        version = 'NONE'
    else:
        version = properties.version

    row = [x+1, docx_file[x], os.path.splitext(docx_file[x])[-1],
                datetime.utcfromtimestamp(Created_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Created_time).strftime('%H:%M:%S'),
                datetime.utcfromtimestamp(Modified_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Modified_time).strftime('%H:%M:%S'),
                Caled_file_size + Unit, Alled_file_size + Unit,
                properties.author, properties.last_modified_by, properties.version,
                hashlib.md5(data).hexdigest(),
                hashlib.sha1(data).hexdigest()]

    with open('DOCX_Metadata.csv', 'a', newline = '') as d_file:
        writer = csv.writer(d_file)
        writer.writerow(row)
    d_file.close()
print('DOCX File Metadata Exporting Finished', '\n')

#"""""""""""""""""""""PPTX FILE METADATA PARSER"""""""""""""""""""""
#Make 'PPTX_Metadata.csv' File
with open('PPTX_Metadata.csv', 'w', encoding='euc-kr', newline = '') as p_file:
    writer = csv.writer(p_file)
    writer.writerow(['Number', "FileName", "FileType", "Created Date", "Created Time(UTC +0:00)",
                "Modified Date", "Modified Time(UTC +0:00)", "File Size", "Disk Allocation Size",
                "Creator", "Last Modified By", "Version", "MD-5", "SHA-1"])

#Choose PPTX Files
pptx_file = []
for x in files:
    if x.endswith('pptx'):
        pptx_file.append(x)
print('PPTX Found:', len(pptx_file), 'Files')
print('DOCX File Metadata Exporting...')
for x in range(0, len(pptx_file)):
    #Collect PPTX File Properties
    Created_time = os.path.getctime(pptx_file[x])
    Modified_time = os.path.getmtime(pptx_file[x])
    p_file = Presentation(str(pptx_file[x]))
    properties = p_file.core_properties

    f = open(pptx_file[x], 'rb')
    data = f.read()
    f.close()

    #Get File Size
    file_size = os.path.getsize(pptx_file[x])
    if file_size >= (1024 * 1024):
        Unit = "MB"
    else:
        Unit = "KB"
    p_bytesize = file_size
    allocation_size = file_size

    #Calculate File Size
    Alled_file_size = str(AllocationSize(allocation_size))[:5]
    #Calculate Disk Alloction size
    Caled_file_size = str(CalculateFileSize(file_size))[:5]

    if properties.version == '':
        version = 'NONE'
    else:
        version = properties.version

    row = [x+1, pptx_file[x], os.path.splitext(pptx_file[x])[-1],
                datetime.utcfromtimestamp(Created_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Created_time).strftime('%H:%M:%S'),
                datetime.utcfromtimestamp(Modified_time).strftime('%Y-%m-%d'),
                datetime.utcfromtimestamp(Modified_time).strftime('%H:%M:%S'),
                Caled_file_size + Unit, Alled_file_size + Unit,
                properties.author, properties.last_modified_by, properties.version,
                hashlib.md5(data).hexdigest(),
                hashlib.sha1(data).hexdigest()]

    with open('PPTX_Metadata.csv', 'a', newline = '') as p_file:
        writer = csv.writer(p_file)
        writer.writerow(row)
    p_file.close()
print('PPTX File Metadata Exporting Finished', '\n')
