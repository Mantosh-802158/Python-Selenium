from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
workbook = xlsxwriter.Workbook('Uni_format2.xlsx')
worksheet = workbook.add_worksheet('sheet2')
merge_format = workbook.add_format(
    {
        "bold":2,
        "fg_color":"grey",
        "align":"center",
        "valign":"vcenter",
    }
)
worksheet.merge_range("C1:Q1","Theory",merge_format)
worksheet.merge_range("R1:W1",'Practical',merge_format)

worksheet.merge_range("C2:E2","English",merge_format)
worksheet.merge_range("F2:H2","CONT",merge_format)
worksheet.merge_range("I2:K2","SAD",merge_format)
worksheet.merge_range("L2:N2","C LANG",merge_format)
worksheet.merge_range("O2:Q2","OS",merge_format)

worksheet.merge_range("R2:T2","C LAB",merge_format)
worksheet.merge_range("U2:W2","OS",merge_format)

worksheet.write('A3','Reg. No')
worksheet.write('B3','Name')
worksheet.write('X3','SGPA')
worksheet.write('Y3','CGPA')
col = 2
def column_name():
    global col
    worksheet.write(2,col,"External")
    col+=1
    worksheet.write(2,col,"Internal")
    col+=1
    worksheet.write(2,col,"Total")
    col+=1
    return 0
for i in range(1,8):
    column_name()

column_data =2
row_data = 3

def write_theory_result(theory_rows):
    global column_data
    global row_data
    for row in range(1,6):
        theory_data = theory_rows[row].find_elements(By.TAG_NAME,'td')
        for data in range(2,5):
            worksheet.write(row_data,column_data,theory_data[data].text)
            column_data+=1

def write_practical_result(practical_rows):
    global column_data
    global row_data
    for row in range(1,3):
        practical_data = practical_rows[row].find_elements(By.TAG_NAME,'td')
        for data in range(2,5):
            worksheet.write(row_data,column_data, practical_data[data].text)
            column_data+=1

def SGPA_result(SGPA):
    global column_data
    global row_data
    worksheet.write(row_data,column_data,SGPA.text)
    column_data+=1
    
def CGPA_result(CGPA):
    global column_data
    global row_data
    worksheet.write(row_data,column_data,CGPA.text)

name_col = 'B'
name_col_var = 4

reg_col = 'A'
reg_col_var = 4

def write_result():
    global column_data
    global row_data
    global reg_col_var
    global name_col_var
    driver = webdriver.Firefox()
    driver.get('https://results.akuexam.net/Vocat2ndSem2023Results.aspx')
    for i in range(22303302001,22303302010):
        column_data=2
        try:
            search = WebDriverWait(driver,20).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_TextBox_RegNo"]'))
            )
            search.send_keys(i)
            search.send_keys(Keys.RETURN)
       
            try:
                
                name = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_DataList1_ctl00_StudentNameLabel"]'))
                )

                
                theory_t_body = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_GridView1"]/tbody'))
                )
                practical_t_body = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_GridView2"]/tbody'))
                )
                SGPA = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_DataList5_ctl00_GROSSTHEORYTOTALLabel"]'))
                )
                CGPA = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_GridView3"]/tbody/tr[2]/td[7]'))
                )

                worksheet.write((reg_col+str(reg_col_var)),i)
                reg_col_var+=1
                worksheet.write((name_col+str(name_col_var)),name.text)
                name_col_var+=1

                theory_rows = theory_t_body.find_elements(By.TAG_NAME,'tr')
                practical_rows = practical_t_body.find_elements(By.TAG_NAME,'tr')
               
                write_theory_result(theory_rows)
                write_practical_result(practical_rows)
                SGPA_result(SGPA)
                CGPA_result(CGPA)
                row_data+=1
               
                print(name.text)
            except:
                print('No such element found ',i)
            finally:
                print('success')
        finally:
            driver.back()
            driver.refresh()

    return 0
    
write_result()
workbook.close()