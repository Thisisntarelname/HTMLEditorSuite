#!/usr/bin/env python
#This takes the xlsx file from /uploads, edits the /Data pages to reflect upadted information from xlsx file. 
import os
import openpyxl
import glob
import json
from bs4 import BeautifulSoup


filenames = []
teachernames = []

for file in glob.glob("uploads/*.xlsx"):
  path = (file)#Needs to open the one file in 'uploads' folder
  print(path)
  delete = path


wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
ws = sheet_obj
 
sheet_object = wb_obj.active
max_rows = sheet_object.max_row#finding how many rows there are

os.chdir('Data')
i = 2#if i=1, it would use "21/22" (the year) as well
while i <= max_rows:
  cell_obj = sheet_obj.cell(row = i, column = 1)
  if (cell_obj.value is None) or (cell_obj.value == "GYM") or (cell_obj.value == "DANCE"):#filters out empty, "dance", and "gym" results
    a = 1
  else:
    txt = cell_obj.value
    x = "RM " in txt
    if (x == True):#filters out room number results
      a = 1
    else:
           
        searchThis = cell_obj.value
        teacherName = searchThis
        searchThis = searchThis.replace(" ", "")

        teachernames.append(teacherName)
        filenames.append(searchThis)

        #this is where we will upadte html with new information
        #open filename and read
        #paste tmplate
        
        
        with open('test.html', 'r') as f:
        
            contents = f.read()
        
            soup = BeautifulSoup(contents, 'lxml')
        
            #---------Subejcts
            tag = (soup.find(id='subjects'))#Finds h3 with id: subjects
            subjects = ["English", "Math", "Science", "Social Studies"]#list of subjects created from xlsx file
            build_subjects = subjects[0]#Starting the loop off with first item
            i = 1
            while i<len(subjects):#couldn't get a for loop to work
              print(i)
              build_subjects = build_subjects+", "+subjects[i]#Concated each item to variable passed form last loop through
              print(build_subjects)
              i = i+1
        
            subject_string = '<h3 id="subject">Subjects: '+build_subjects+'</h3>'
            print(subject_string)
            tag.replace_with(subject_string)
            #---------Classes
            #---------Rooms
        
          
          
        with open("example_modified.html", "w") as f_output:
        
            f_output.write(soup.prettify(formatter=None)) 
        
        



  i = i+1
  

print("filename", filenames)
print("teachernames", teachernames)


