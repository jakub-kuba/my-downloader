# Author: jakub-kuba

import time
from datetime import datetime
from selenium import webdriver
import threading
import schedule
import os
import sys
import win32com.client
import pythoncom

def get_element(driver, seconds, link):
    """Finds HTML element"""
    counter = 0
    while counter < seconds:
        try:
            return driver.find_element_by_xpath(link)
        except:
            counter += 1
            time.sleep(1)
    else:
        sec = 3
        while sec != 0:
            try:
                print("Element not found. Next attempt will be made in 60 seconds.")
                time.sleep(60)
                return driver.find_element_by_xpath(link)
            except:
                sec -= 1
                time.sleep(1)
        else:
            print("Element not found!")


def get_website(driver, seconds, link):
    """Opens required website"""
    counter = 0
    while counter < seconds:
        try:
            return driver.get(link)
        except:
            counter += 1
            time.sleep(1)
    else:
        sec = 2
        while sec != 0:
            try:
                print("Website not found. Next attempt will be made in 3 minutes.")
                time.sleep(180)
                return driver.get(link)
            except:
                sec -= 1
                time.sleep(1)
        else:
            print("Website not found!")


def get_xpath_click(driver, seconds, link):
    """Finds HTML element by xpath and clicks"""
    counter = 0
    while counter < seconds:
        try:
            element = driver.find_element_by_xpath(link)
            return element.click()
        except:
            counter += 1
            time.sleep(1)
    else:
        print("Element not found")


def driver_options(chromedriver, destination, argument="s"):
    """Sets webdriver options"""
    prefs = {"download.default_directory": destination}
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option("prefs", prefs)
    if argument == 'h':
        options.add_argument("--headless=new")
    driver = webdriver.Chrome(chromedriver, options=options)
    return driver


def check_chromedriver(chromedriver, destination):
    """Checks if chromedriver version is up tp date"""
    try:
        driver = driver_options(chromedriver, destination, "h")
        driver.close()
        print("\nThe Chromedriver version is up to date.\n")
    except:
        print("Please update Chromedriver.exe and restart the program\n!")
        sys.exit()


def count_and_finish(destination, rep_name, ext, time_limit, driver):
    """Counts files with specific name and changes name of downloaded file"""
    list_of_files = os.listdir(destination)
    list_limited = [x for x in list_of_files if rep_name in x and 'crdownload' not in x and x[0] != '~']
    first_list = list_limited.copy()
    len_limited = len(list_limited)
    len_limited_plus = len_limited + 1
    counter = 0

    while len_limited < len_limited_plus:
        list_of_files = os.listdir(destination)
        list_limited = [x for x in list_of_files if rep_name in x and 'crdownload' not in x and x[0] != '~']
        len_limited = len(list_limited)
        time.sleep(1)
        counter += 1
        #seconds counter
        print(f'Download in progress: {counter} / {time_limit} [s]', end="\r", flush=True)
        if counter > time_limit:
            print("Download time too long. The file has not been downloaded!\n")
            time.sleep(1)
            status = 'ERROR'
            driver.close()
            email_send(rep_name, status, destination)
            break

    else:
        day = datetime.now().strftime('%d-%m-%Y_%H-%M')
        file_downloaded = [x for x in list_limited if x not in first_list][0]
        final_name = rep_name+" "+day+ext
        #rename the downloaded file
        try:
            os.rename('my downloads/'+file_downloaded, 'my downloads/'+final_name)
        except:
            print("except")
            return None
        dowload_time = seconds_to_minutes(counter)
        print("The file has been downloaded! Download time:", dowload_time, "\n")
        status = "DOWNLOADED"
        email_send(rep_name, status, destination, dowload_time)
        time.sleep(1)
        driver.close()


def my_log():
    """Adds date and time in the text file"""
    now = datetime.now()
    date = now.strftime('%d-%b-%Y_%H_%M_%S')
    with open("logfile.txt", "a") as myfile:
        myfile.write(date+"\n")


def seconds_to_minutes(seconds):
    """Converts seconds to minutes"""
    minutes = (seconds % 3600) // 60
    seconds = (seconds % 3600) % 60
    return f'{minutes:02} min {seconds:02} sec'


def email_send(rep_name, status, destination, download_time=None):
    """Sends email with download status"""
    outlook = win32com.client.Dispatch('outlook.application', pythoncom.CoInitialize())
    mail = outlook.CreateItem(0)
    mail.To = 'someone999@exampledomain.com'
    mail.Subject = f'My Downloader - File Status: {status}'
    now = datetime.now()
    date = now.strftime('%d-%b-%Y %H:%M:%S')
    this_time = date
    br = '</br><br>'
    location = f'<a href="{destination}">my downloads folder</a>'
    if status == 'DOWNLOADED':
        mail.HTMLBody = ("file name:"+br+"<b>"+rep_name+"</b>"+br+br+
                         "status:"+br+"<b>"+status+"</b>"+br+br+
                         "location:"+br+location+br+br+
                         "date and time (CET)"+br+this_time+br+br+
                         "download time:"+br+"<b>"+download_time+"</b>")
    else:
        mail.HTMLBody = ("file name:"+br+"<b>"+rep_name+"</b>"+br+br+
                         "status:"+br+"<b>"+status+"</b>"+br+br+
                         "date and time (CET)"+br+this_time)
    mail.Display()


def run_threaded(job_func):
    """Function needed for multithreading"""
    job_thread = threading.Thread(target=job_func)
    job_thread.start()


def main():

    # each function below is for one file to be downloaded
    # address - website address
    # rep_name - file name
    # ext - file extension
    # time_limit - max time(seconds) for downloading the file

    def records_data():
        address = 'https://www.appsloveworld.com/sample-excel-data-for-analysis?utm_content=cmp-true'
        rep_name = "Records_Data"
        ext = '.xlsx'
        time_limit = 50
        continue_element = '//*[@id="ez-accept-all"]'
        close_ad = '//*[@id="ezmob-footer-close"]'
        download_element = '//*[@id="divcontent"]/div[2]/div[2]/div[14]/div[2]/a'

        print(datetime.now().strftime('%H:%M'), " - starting:", rep_name)
        driver = driver_options(chromedriver, destination, "h")

        get_website(driver, 10, address)
        get_xpath_click(driver, 10, continue_element)
        get_xpath_click(driver, 10, close_ad)
        get_xpath_click(driver, 10, download_element)
        count_and_finish(destination, rep_name, ext, time_limit, driver)


    def fruit_market():
        address = 'https://data.jakub-kuba.com/market'
        rep_name = 'fruit'
        ext = '.csv'
        time_limit = 10
        data_element = '/html/body/div[1]/div[1]/a[2]/div/img'
        download_element = '/html/body/div[1]/div[2]/div[1]/li/a/b'

        print(datetime.now().strftime('%H:%M'), " - starting:", rep_name) 
        driver = driver_options(chromedriver, destination, "h")

        get_website(driver, 10, address)
        get_element(driver, 10, data_element)
        get_xpath_click(driver, 10, download_element)
        count_and_finish(destination, rep_name, ext, time_limit, driver)

    #subfolder which stores downloaded files
    dest_folder = "\my downloads"
    destination = os.getcwd()+dest_folder

    #chromedriver location
    chromedriver = r'C:\\chromedriver\chromedriver.exe'

    #check if chromedriver needs to be updated
    check_chromedriver(chromedriver, destination)

    print("The program is active. Do not close the terminal.\n")

    #add date & time every minute
    schedule.every(1).minutes.do(run_threaded, my_log)

    #the files are downloaded according to the schedule below
    schedule.every().day.at("10:00").do(run_threaded, records_data)
    schedule.every().wednesday.at("11:30").do(run_threaded, fruit_market)

    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__== "__main__":
    main()