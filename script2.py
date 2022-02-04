#This is to extract all the data from the list of URLs

#Import Dependencies
from openpyxl import Workbook, load_workbook
import requests
from bs4 import BeautifulSoup

ownership_type, num_students = [],[]

#Put all the URLs in a list
curl_list = []

wb_read = load_workbook("URL File.xlsx")
ws_read = wb_read["Colleges"]

#Read a range of values from a particular column
while True:
    try:
        start_num = int(input("Enter the starting number: "))
        end_num = int(input("Enter the ending number: "))
        col_name = "e" #input("Enter column name: ")
    except:
        print("Enter valid values")
    if len(col_name)>1:
        continue
    if col_name.isalpha():
        col_name = col_name.upper()
    else:
        continue
    if (start_num <= 0) or (end_num <= 0):
        print("Enter positive integers only")
        continue
    if (start_num >= end_num):
        print("Starting number cannot be greater than ending number")
        continue
    break

for i in range(start_num, end_num+1):
    cell_name = col_name + str(i)
    curl_list.append(ws_read[cell_name].value)
    

for i in range(0,len(curl_list)):
    if curl_list[i].startswith("https://www.google"):
        ownership_type.append("NA")
        num_students.append("NA")
    else:
        response = requests.get(curl_list[i])
        soup = BeautifulSoup(response.text, 'lxml')
        try:
            printnow = soup.find('div', class_ = "cardBlkInn quickFact").text.replace("\t","").strip()
            ownership_start = printnow.find("Ownership")
            student_num_start = printnow.find("Total Student Enrollments")
            ownership_type.append(printnow[ownership_start+11:ownership_start+18])
            num_students.append(printnow[student_num_start+26:].strip())
        except:
            ownership_type.append("NA")
            num_students.append("NA")

#Write the extracted data into another column of the Excel file
while True:
    ownership_col = "b" #input("Column to enter the ownership details: ")
    students_col = "c" #input("Column to enter the number of students: ")
    if (len(ownership_col) > 1) or (len(students_col) > 1):
        continue
    if (ownership_col.isalpha()) and (students_col.isalpha()):
        ownership_col = ownership_col.upper()
        students_col = students_col.upper()
    if (ownership_col == col_name) or (students_col == col_name):
        print("Cannot override the name of the college")
        continue
    break
    
start_num_dup = start_num
start_num_dup1 = start_num
#end_num_dup = end_num
for i in ownership_type:
    cell_name = ownership_col+str(start_num_dup)
    start_num_dup += 1
    ws_read[cell_name].value = i

for i in num_students:
    cell_name = students_col+str(start_num_dup1)
    start_num_dup1 += 1
    ws_read[cell_name].value = i
    
wb_read.save("Details.xlsx")