from selenium.webdriver import Firefox
from selenium.webdriver import FirefoxProfile
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as expected
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from time import sleep
import os
import os.path
from glob import glob
import xlrd
import sqlite3
import datetime

profile = FirefoxProfile()
profile.set_preference("browser.download.panel.shown", False)
profile.set_preference("browser.helperApps.neverAsk.openFile","text/xls,application/vnd.ms-excel")
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/xls,application/vnd.ms-excel")
profile.set_preference("browser.download.folderList", 2);
profile.set_preference("browser.download.dir", os.path.join(os.getcwd(),'excel'))

options = Options()
options.add_argument("--headless")
driver = Firefox(firefox_options=options,firefox_profile=profile)
wait = WebDriverWait(driver, timeout=10)

print("Firefox Headless Browser Invoked")
driver.get('https://www.xmarkets.db.com/DE/ENG/Product_Overview/WAVEs_Call')
print("Got web page")

Select(wait.until(expected.visibility_of_element_located((By.CSS_SELECTOR, '#ctl00_leftNavigation_ctl00_asidebar-tab-navigation_DerivativeFilter_UnderlyingTypeFilterDropdown')))).select_by_visible_text("Indices Germany")
sleep(2)
Select(wait.until(expected.visibility_of_element_located((By.CSS_SELECTOR, '#ctl00_leftNavigation_ctl00_asidebar-tab-navigation_DerivativeFilter_UnderlyingFilterDropdown')))).select_by_visible_text("DAX")
sleep(2)
driver.get(driver.current_url)
sleep(5)
driver.execute_script('''
window.downloadTheExcel =  function(){
  elem = $('a[href^="javascript:CreateSelectionExcel"]')[0];
  urltodownload = elem.href;
  command = urltodownload.slice('javascript:'.length,urltodownload.length);
  eval(command)
}
''')

def downloadTheExcel():
    driver.execute_script('downloadTheExcel()')

database = 'main.db'

newDatabase = not os.path.isfile('main.db')

if newDatabase:
    conn = sqlite3.connect('main.db')
    c = conn.cursor()
    c.execute('''DROP TABLE IF EXISTS stocks''')
    c.execute('''CREATE TABLE stocks
                 (
                 name text,
                 wkn text,
                 bid number,
                 ask number,
                 strike_price number,
                 barrier_level number,
                 ref number,
                 distance_to_barrier_level number,
                 ratio number,
                 leverage number,
                 maturity text,
                 date text
                 )''')
    conn.commit()
    conn.close()

def save():
    conn = sqlite3.connect('main.db')
    c = conn.cursor()
    files = glob('excel/*')
    for file in files:
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        date = sheet.cell_value(rowx=7, colx=6)
        date = datetime.datetime(*xlrd.xldate_as_tuple(date, workbook.datemode))
        for rowidx in range(10,sheet.nrows):
            row = sheet.row(rowidx)
            c.execute('INSERT INTO stocks values (?,?,?,?,?,?,?,?,?,?,?,?)',
                      (row[0].value,
                       row[1].value,
                       row[2].value,
                       row[3].value,
                       row[4].value,
                       row[5].value,
                       row[6].value,
                       row[7].value,
                       row[8].value,
                       row[9].value,
                       row[10].value,
                       date))
        conn.commit()
        workbook.release_resources()
        del workbook
        os.remove(file)
    conn.close()


from apscheduler.schedulers.background import BlockingScheduler
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.combining import OrTrigger
from apscheduler.triggers.combining import AndTrigger

sched = BlockingScheduler()

every15Minute = CronTrigger(minute='0,15,30,45')

between = CronTrigger(hour='7-22')



sched.add_job(downloadTheExcel, every15Minute)

cleaner = BackgroundScheduler()
everyHour = CronTrigger(minute='7')

cleaner.add_job(save,everyHour)

cleaner.start()
sched.start()
