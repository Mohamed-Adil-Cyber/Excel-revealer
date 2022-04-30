import argparse
from pathlib import Path
import os
import shutil
from os import path
from shutil import make_archive
from zipfile import ZipFile
from pathlib import Path
import sys
from os import walk



#creating all cmd arguments
parser = argparse.ArgumentParser()


#input file argument
parser.add_argument('-i', '--input',type=str, required=True, 
    help="excel input file")


#output file argument
parser.add_argument('-o', '--output', type=str, required=True, 
    help="excel output file name")

#argument for choosing which option
parser.add_argument('-c','--choice', dest='format', choices=['reveal', 'unprotect'], required=True, 
                    help="shows datetime in given format")

#parsing all arguments
args = parser.parse_args()

#option variable
choice = args.format


#Code for revealing hidden sheets
if choice == 'reveal':
    
    #name of the excel and the output excel
    inputfile = args.input
    outputfile = args.output

    #extracting excel objects
    with ZipFile(inputfile , 'r') as zip_ref:
        zip_ref.extractall("test")

    #files focused on in editing    
    infile = "test/xl/workbook.xml"
    outfile = "test/xl/test.xml"

    #Editing the xml files
    delete_list = ["veryHidden", "Hidden"]
    try:
           with open(infile) as fin, open(outfile, "w+") as fout:
                for line in fin:
                     for word in delete_list:
                        line = line.replace(word, "")
                     fout.write(line)
    except:
           pass
    

    #replacing unedited file with the new file         
    os.remove(infile)
    os.rename(outfile, infile)

    #Making the excel output file
    shutil.make_archive(outputfile, 'zip', "/test/")
    pre, ext = os.path.splitext('/result.zip')
    os.rename(outputfile+'.zip', outputfile + ".xlsx")
    shutil.rmtree("/test")







#Code for erasing sheet protection    
elif choice == 'unprotect':

    #name of the excel and the output excel
    inputfile = args.input
    outputfile = args.output

    #extracting excel objects
    with ZipFile(inputfile , 'r') as zip_ref:
        zip_ref.extractall("test")

        
    #Listing all sheets in the file    
    filenames = next(walk('test/xl/worksheets'), (None, None, []))[2]

    #Iterating through all sheets
    for sheet in filenames:
        with open('test/xl/worksheets'+'/'+sheet,'r') as file:
            
          #Searching for sheet protection
          data = file.read()
          text = ''
          try:
              text = data[data.index("<sheetProtection "):data.index("scenarios=")+15]
          except:
              pass

        #files focused on in editing     
        infile = "test/xl/worksheets" +'/'+sheet
        outfile = "test/xl/worksheets/test.xml"

        
        #Editing the sheets
        delete_list = [text,"veryHidden", "Hidden"]
        with open(infile) as fin, open(outfile, "w+") as fout:
            for line in fin:
                 for word in delete_list:
                    line = line.replace(word, "")
                 fout.write(line)



        #replacing unedited file with the new file          
        os.remove(infile)
        os.rename(outfile, infile)

    #Making the excel output file    
    shutil.make_archive(outputfile, 'zip', "test/")
    pre, ext = os.path.splitext('/result.zip')
    os.rename(outputfile+'.zip', outputfile + ".xlsx")
    shutil.rmtree("test")


