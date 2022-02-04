#This is to collect all the URLs first

#Import dependencies
from openpyxl import Workbook, load_workbook
from selenium import webdriver

#Read contents from existing workbook
wb_read = load_workbook("URL File.xlsx")
ws_read = wb_read["Colleges"]

#Read a range of values from a particular column
while True:
    try:
        start_num = int(input("Enter the starting number: "))
        end_num = int(input("Enter the ending number: "))
        col_name = "a"#input("Enter column name: ")
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
    j = ""
    for k in str(query_list[i]):
        if k.isnumeric():
            break
        elif k in ".-:":
            continue
        elif k == ",":
            j+= " "
        else:
            j += k
    j = j[:-3]
    j = j.replace("  "," ")
    j = j.replace(" ","+")
    search_query = 'https://www.google.com/search?q='+j+'+"overview"+careers360&start=1'
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

#Process curl_list now
curl_list_dup = []
for i in curl_list:
    j = ""
    count = 0
    for k in range(0,len(i)):
        if i[k] == '/':
            count += 1
            if count == 5:
                break
            else:
                j += i[k]
        else:
            j += i[k]
    curl_list_dup.append(j)



#Write all the URLs in the Excel File
url_col = "E" #input("Enter the column for the URLs: ")
start_num_dup = start_num
start_num_dup1 = start_num
#end_num_dup = end_num
for i in curl_list_dup:
    cell_name = url_col+str(start_num_dup)
    start_num_dup += 1
    ws_read[cell_name].value = i
    
wb_read.save("URL File.xlsx")