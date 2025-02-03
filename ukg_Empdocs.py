from selenium import webdriver #Web driver activities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from configparser import ConfigParser #Configuration read
from pathlib import Path #Path conversions
from os import listdir
from os.path import isfile, join
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
# import undetected_chromedriver as uc
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
import pyodbc
import logging #Activity logging
import time #Sleep
import pyautogui
import time
import pygetwindow
import shutil
import keyboard   #from webdriver_manager.chrome import ChromeDriverManager
 
 
download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
firefox_location='./'
 
 
global wait, driver, webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection,min
 
 
def read_config_file():
    global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
    global sql_server_name, sql_user_name, sql_password, sql_db
    try:
        # configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        # webpage_url = config.get ("Data", "webpage_URL")
        # user_name = config.get ("Data", "user_name")
        # password = config.get ("Data", "password")
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as e:
        print('config file read error')
        return(-1)
 
#--------------------------------------------------------------------------------------------------------
 
 
def setup():
    global firefox_location, download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/zip")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "image/jpeg")
    profile.set_preference("browser.helperApps.saveToDisk.image/jpeg", "application/octet-stream")
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.viewableInternally.enabledTypes", "")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("browser.download.manager.useWindow", False)
    profile.set_preference("browser.download.manager.closeWhenDone", True)
    profile.set_preference("print_printer", "Microsoft Print to PDF")
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("print.show_print_progress", False)
    profile.set_preference("print.save_as_pdf.links.enabled", True)
    profile.set_preference("browser.helperApps.alwaysAsk.force", False)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    # driver = webdriver.Firefox(service_log_path='NUL', firefox_profile=profile)
    driver = webdriver.Firefox()
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    webpage_url = "https://ew33.ultipro.com"
    driver.get(webpage_url)
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    return 0
 
#--------------------------------------------------------------------------------------------------------
 

def addToDatabase(src_id,emp_name,document_title,category):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [Orchestra].[dbo].[Emp_Files_emp_docs] ([Employee_Number],[Emp_Name],[Category],[File_Name],[IsDownloaded]
                ) VALUES (?, ?, ?, ?, ?)
            """
            values = (src_id,emp_name,category,document_title, 1)
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(src_id," :Added to Emp_Files")

    except Exception as d:
        print('Database error which is ', d)
 
#--------------------------------------------------------------------------------------------------------
 
def remove_files():
    files_to_delete = os.listdir(download_dir)
    for filename in files_to_delete:
        file_path = os.path.join(download_dir, filename)
        os.remove(file_path)  # Delete the file
 
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)
 
#--------------------------------------------------------------------------------------------------------
 
def login():
    global min
    #username input xpath
    # Find and interact with username field
    s=input("enter_before:")
    xpath_username = '//*[@id="ctl00_Content_Login1_UserName"]'
    element_username = driver.find_element(By.XPATH, xpath_username)
    element_username.click()
    element_username.clear()
    element_username.send_keys("TAP.SMC@tapinnov.com")

    time.sleep(2)
 
    # Find and interact with password field
    xpath_password = '//*[@id="ctl00_Content_Login1_Password"]'
    element_password = driver.find_element(By.XPATH, xpath_password)
    element_password.click()
    element_password.clear()
    element_password.send_keys("Welcome@2025")
    time.sleep(2)

    # Find and click login button
    xpath_login_button = '//*[@id="ctl00_Content_Login1_LoginButton"]'
    element_login_button = driver.find_element(By.XPATH, xpath_login_button)
    element_login_button.click()
    # time.sleep(10)
    s= input("enter_after:")
 
    #Employee symbol
    xpath = '//*[@id="menu_admin"]'
    element = driver.find_element(By.XPATH,xpath)
    element.click()
    time.sleep(5)
    #My Employees
    xpath = '//*[@id="424"]'
    element = driver.find_element(By.XPATH,xpath)
    element.click()
    time.sleep(5)
   
    element = driver.find_element(By.ID,"ContentFrame")
    driver.switch_to.frame(element)
    element = driver.find_element(By.XPATH,'//*[@id="GridView1_firstSelect_0"]')
    select = Select(element)
    element.click()
    element.send_keys("Employee Number")
    time.sleep(2)
    element = driver.find_element(By.XPATH,'//*[@id="GridView1_Operator_0"]')
    element.send_keys("is")
    time.sleep(2)
 
#--------------------------------------------------------------------------------------------------------
 
# def createFolder(sFolder):
#     isExist = os.path.exists(sFolder)
#     if not isExist:
#         os.makedirs(sFolder)
#     return(0)
 
def connectSQL():
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as e:
        print('SQL connection error')
        return(-1)
    
def sanitize_filename(filename):
    """Remove or replace invalid characters in filenames."""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)  # Replace special characters with '_'

def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir, src_id, emp_name):
    try:
        # Sanitize the document title
        sanitized_title = sanitize_filename(document_title)

        # Define new folder path in the target directory based on category
        new_folder = os.path.join(target_dir, category)
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)

        # Define the base new file path
        base_new_file_path = os.path.join(new_folder, sanitized_title + file_extension)

        # Check if the file already exists and generate a new name if necessary
        new_file_path = base_new_file_path
        counter = 1
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(new_folder, f"{sanitized_title}_{counter}{file_extension}")
            counter += 1

        # Rename and move the file
        original_file = os.path.join(download_dir, file_name)
        shutil.move(original_file, new_file_path)
        print(f"Moved file '{file_name}' to '{new_file_path}'")

        # Log the downloaded file
        value = f'"{fld_name}","{sanitized_title}"'
        with open('Downloaded_files.csv', "a") as f:
            f.write(value + '\n')

        addToDatabase(src_id, emp_name, sanitized_title, category)

    except Exception as e:
        print(f"Error while renaming and moving the file: {e}")

def wait_for_download(download_dir, timeout=30):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = any([filename.endswith('.part') for filename in os.listdir(download_dir)])
        seconds += 1
    return not dl_wait

# def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir, src_id, emp_name):
#     try:
#         # Define new folder path in the target directory based on category
#         new_folder = os.path.join(target_dir, category)
#         if not os.path.exists(new_folder):
#             os.makedirs(new_folder)

#         # Define the base new file path
#         base_new_file_path = os.path.join(new_folder, document_title + file_extension)
        
#         # Check if the file already exists and generate a new name if necessary
#         new_file_path = base_new_file_path
#         counter = 1
#         while os.path.exists(new_file_path):
#             new_file_path = os.path.join(new_folder, f"{document_title}_{counter}{file_extension}")
#             counter += 1

#         # Rename and move the file
#         original_file = os.path.join(download_dir, file_name)
#         shutil.move(original_file, new_file_path)
#         print(f"Moved file '{file_name}' to '{new_file_path}'")

#     # except Exception as e:
#     #     print(f"Error: {e}")
#         value = '"'+fld_name+'"'+","+'"'+document_title+'"'
#         f= open('Downloaded_files.csv',"a")
#         f.write(value)
#         f.write('\n')
#         f.close()
#         addToDatabase(src_id,emp_name,document_title,category)
#     except Exception as e:
#         print(f"Error while renaming and moving the file: {e}")
 
# def wait_for_download(download_dir, timeout=30):
#     seconds = 0
#     dl_wait = True
#     while dl_wait and seconds < timeout:
#         time.sleep(1)
#         dl_wait = any([filename.endswith('.part') for filename in os.listdir(download_dir)])
#         seconds += 1
#     return not dl_wait
   
#--------------------------------------------------------------------------------------------------------
 
def searchanddownload(fld_name, src_id, emp_name):
    global SQLconnection,min
    try:
        print("src Id ", src_id)
        driver.refresh()
        window_handles = driver.window_handles
        print("a ", len(window_handles))
        a = len(window_handles)
        b = 0
        while (a>1):
            driver.switch_to.window(window_handles[b-1])
            driver.close()
            b -=1
            a-=1
        window_handles = driver.window_handles
        driver.switch_to.window(window_handles[0])
        time.sleep(5)
        try:
            element = driver.find_element(By.ID,"ContentFrame")
            driver.switch_to.frame(element)
            time.sleep(3)
            print("switched")
        except:
            print("already switched")

        element = driver.find_element(By.XPATH,'//*[@id="GridView1_firstSelect_0"]')
        select = Select(element)
        element.click()
        element.send_keys("Employee Number")
        time.sleep(2)
        element = driver.find_element(By.XPATH,'//*[@id="GridView1_Operator_0"]')
        element.send_keys("is")
        time.sleep(2)
 
        input_element = driver.find_element(By.ID, 'GridView1_TextEntryFilterControlInputBox_0')
        input_element.click()
        input_element.clear()
        input_element.send_keys(src_id)
        time.sleep(3)
        keyboard.press_and_release('enter')
        time.sleep(4)
        # driver.refresh()
       
        try:
            element = driver.find_element(By.ID,"ContentFrame")
            driver.switch_to.frame(element)
        except:
            print('already in content frame \n')

        element = driver.find_element(By.XPATH, '//*[@id="ctl00_Content_GridView1"]/tbody/tr/td[3]')
        print(element.text)
        time.sleep(5)

        time.sleep(3)
        print(f"\n\n Elemetne TExt = {element.text} | Source id = {src_id}")
 
        if element.text == src_id:
            #First employee
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Content_GridView1"]/tbody/tr/td[1]/a'))
            )
            element.send_keys(Keys.ENTER)
       
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            time.sleep(4)
            driver.maximize_window()
            # s = len(window_handles)
            # while (s>2):
            #     driver.switch_to.window(window_handles[b-1])
            #     driver.close()
       
            #Employee Documents
            xpath = '//*[@id="1305"]'
            element = driver.find_element(By.XPATH,xpath)
            element.click()
            time.sleep(10)
 
            try:
                element = driver.find_element(By.ID, "ContentFrame")
                driver.switch_to.frame(element)
            except:
                print("No content frame found")
            time.sleep(3)
       
            table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_Content_gvFileInfo"))
            )
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]
            document_count = len(rows)
            print(f"Total documents: {document_count}")
            with open('document_count.csv', 'a', newline='') as count_file:
                count_writer = csv.writer(count_file)
                count_writer.writerow(["Folder name", "Document Count"])
                count_writer.writerow([fld_name,document_count])
            print(f"Total documents: {document_count} (saved to 'document_count.csv')")
 
            # if document_count > 10:
            try:
                dropdown = Select(driver.find_element(By.ID,'gvFileInfo_numRowsSelect')) # Select the option with value '50'
                dropdown.select_by_value('50')
            except Exception as e:
                print("not clicking value 50")
                # dropdown = driver.find_element(By.XPATH,'//*[@id="gvFileInfo_numRowsSelect"]')
                # dropdown.click()
                # time.sleep(2)
                # value = driver.find_element(By.XPATH,'//*[@id="gvFileInfo_numRowsSelect"]/option[3]')
                # value.click()

            # elif document_count > 10:
            #     try:
            #         dropdown = Select(driver.find_element(By.ID,'gvFileInfo_numRowsSelect')) # Select the option with value '50'
            #         dropdown.select_by_value('50')
            #     except Exception as e:
            #         print("not clicking value 50")
 
            try:
                table = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ctl00_Content_gvFileInfo"))
                )
                rows = table.find_elements(By.TAG_NAME, "tr")[1:]
                with open('Files_in_system.csv', 'a', newline='') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerow(["Folder name", "Document Title", "Category", "Date Added"])
                    for row in rows:
                        columns = row.find_elements(By.TAG_NAME, "td")
                        document_title = columns[1].text
                        category = columns[2].text
                        date_added = columns[4].text
                        csv_writer.writerow([fld_name, document_title, category, date_added])
       
                print("CSV file 'output.csv' has been generated successfully.")
       
            except Exception as e:
                print("Error:", e)
 
            download_dir = r'C:\Users\RPATEAMADMIN\Downloads'  # Change this to your download directory path
            # download_dir = r'C:\Users\rmada\Documents\Jovia\sample\temp_downloads'
 
            target_dir = "V:/Navya/downloads/surlean/Active/"+fld_name+"/"
 
             
            # Locate all rows in the table
            rows = driver.find_elements(By.XPATH, "//tr[@class='GridRowStyle' or @class='altShading']")
            print(f"Found {len(rows)} rows.")
            count = str(len(rows))
            value = '"'+fld_name+'"'+","+count
            f= open('Report_count.csv',"a")
            f.write(value)
            f.write('\n')
            f.close()

            # for i in range(0,4):
                # try:
                #     xpath = '//*[@id="grvPRPayRegister_pageEntry"]'
                #     element = driver.find_element(By.XPATH,xpath)
                #     element.clear()
                #     element.send_keys(i)
                #     time.sleep(1)
                #     keyboard.press_and_release('enter')
                #     time.sleep(3)
                # except:
                #     # j = False
                #     print("loop exist")
           
            for i, row in enumerate(rows):
                try:
                    # Get the document title, category, and file type
                    document_title = row.find_element(By.XPATH, ".//span[contains(@id, 'clDocumentTitle')]").text
                    category = row.find_element(By.XPATH, ".//span[contains(@id, 'cdlCategory')]").text
            
                    # Determine file type (PDF or JPG)
                    # file_type_element = row.find_element(By.XPATH, ".//img[contains(@src, 'file_pdf.png') or contains(@src, 'file_jpg.png')]")
                    # file_extension = ".pdf" if 'file_pdf.png' in file_type_element.get_attribute('src') else ".jpg"
                    file_type_element = row.find_element(By.XPATH, ".//img[contains(@src, 'file_pdf.png') or contains(@src, 'file_jpg.png') or contains(@src, 'file_png.png') or contains(@src, 'file_docx.png') or contains(@src, 'file_xlsx.png') or contains(@src, 'file_txt.png')]")

                    src_attribute = file_type_element.get_attribute('src')
                    if 'file_pdf.png' in src_attribute:
                        file_extension = ".pdf"
                    elif 'file_jpg.png' in src_attribute:
                        file_extension = ".jpg"
                    elif 'file_png.png' in src_attribute:
                        file_extension = ".png"
                    elif 'file_docx.png' in src_attribute:
                        file_extension = ".docx"
                    elif 'file_xlsx.png' in src_attribute:
                        file_extension = ".xlsx"
                    elif 'file_txt.png' in src_attribute:
                        file_extension = ".txt"
                    else:
                        # file_extension = "" 
                        print("none of the type")
        
                    # List files in the download directory before download
                    before_download = set(os.listdir(download_dir))
        
                    # Click the symbol to download the document
                    file_link = row.find_element(By.XPATH, ".//a[contains(@href, 'ViewFileContent.aspx')]")
                    file_link.click()
                    time.sleep(4)
        
                    # Wait for the download to complete
                    if wait_for_download(download_dir):
                        # List files in the download directory after download
                        after_download = set(os.listdir(download_dir))
                        
                        new_files = after_download - before_download
                        if new_files:
                            new_file = new_files.pop()
                            rename_and_move_file(new_file, document_title, category, file_extension, target_dir, fld_name, download_dir,src_id,emp_name)
                            time.sleep(1)
                    else:
                        print(f"Download timed out for row {i + 1}.")
        
                except Exception as e:
                    print(f"Error processing row {i + 1}: {e}")
                    error_row = str(i+1)
                    value = '"'+fld_name+'"'+","+error_row
                    f= open('Error_files.csv',"a")
                    f.write(value)
                    f.write('\n')
                    f.close()

    
        return 0
 
    except Exception as e:
        print("Employee documents Not downloaded\n")
        f= open('Not_downloaded.csv',"a")
        f.write(fld_name)
        f.write('\n')
        f.close()
        print("error ",e)
        return -1
 
#--------------------------------------------------------------------------------------------------------
 
 
def main():
    global SQLconnection
    setup()
    read_config_file()
    connectSQL() 
    login()
 
    import csv
    rows = []
    # with open("Active_roster.csv", 'r') as file:
    #     csvreader = csv.reader(file)
    #     #header = next(csvreader)
    #     for row in csvreader:
    #         rows.append(row)
    # print(len(rows))
    rows = []
    
    select_statement = "SELECT [Emp_Name],[Src_ID],[IsDownload]"
    select_statement = select_statement + "FROM [Surlean].[dbo].[Emp_List]"
    # select_statement += "WHERE [Company] = 'BerlinRosen';"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    rows = cursor.fetchall()
    cursor.close()

    for i in range(0,500):
        each = rows[i]
        src_id = each[1].strip()
        print(src_id)
        # last_name = each[1].strip()
        # first_name = each[2].strip()
        emp_name = each[0].strip()
        # emp_name = last_name+", "+first_name
        print(emp_name)
        path_1 = r'C:\Users\RPATEAMADMIN\Downloads'
        
        files = os.listdir(path_1)
        # Iterate over the files and delete them
        for file in files:
            file_path = os.path.join(path_1, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted {file_path}")
        path = "V:/Navya/downloads/surlean/Active/"
        directory_contents = os.listdir(path)
        fld_name = emp_name+" ("+src_id+")"
        if fld_name not in directory_contents:
            print("entered_searchanddownload")
            # err_f = 1
            # count_flag = 1
            # while(err_f):
            res = searchanddownload(fld_name, src_id, emp_name)
            print("res=",res)
            if res == 0:
                cursor = SQLconnection.cursor()
                cursor.execute("UPDATE [Surlean].[dbo].[Emp_List] SET [IsDownload] = ? WHERE [Src_ID] = ?;",
                (1, src_id))
                print(src_id," :Updated in Emp_List")
                SQLconnection.commit()


            elif res<0:
                print('Employee error '+src_id+" "+emp_name)
                continue

 
#--------------------------------------------------------------------------------------------------------
 
main()