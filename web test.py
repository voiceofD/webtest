#Import dependencies
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from bs4 import BeautifulSoup
import requests

#Read contents from existing workbook
wb_read = load_workbook("Data File.xlsx")
ws_read = wb_read["Colleges"]

#Read a range of values from a particular column
while True:
    try:
        start_num = int(input("Enter the starting number: "))
        end_num = int(input("Enter the ending number: "))
        col_name = input("Enter column name: ")
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

query_list = []
for i in range(start_num, end_num+1):
    cell_name = col_name + str(i)
    query_list.append(ws_read[cell_name].value)


#Create the web driver
driver = webdriver.Chrome()
for i in range(0,len(query_list)):
    j = str(query_list[i]).replace(" ","+")
    search_query = "https://www.google.com/search?q="+j+"+careers360&start=1"
    driver.get(search_query)
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[i+1])
    
curl_list = []

while True:
    waiting = input("Ready for collection? (y/n): ")
    if waiting == "y":
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            curl_list.append(driver.current_url)
        break
    else:
        continue

curl_list = curl_list[:-1]

print(curl_list)

ownership_type, num_students = [],[]
#Get the HTML code from the websites that needs to be scraped
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

print(ownership_type)
print(num_students)

#Write the extracted data into another column of the Excel file
while True:
    ownership_col = input("Column to enter the ownership details: ")
    students_col = input("Column to enter the number of students: ")
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
    
wb_read.save("Data File.xlsx")