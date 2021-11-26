# %%
from selenium import webdriver;
from selenium.webdriver.common.by import By;
from selenium.webdriver.common.keys import Keys;
from selenium.webdriver.common.action_chains import ActionChains;
from webdriver_manager.chrome import ChromeDriverManager;
import datetime;
from time import sleep;
from selenium.webdriver.support.ui import Select
import win32api;
import sys;

# %%
# Initialize Chrome
driver = webdriver.Chrome();
driver.maximize_window();

# %%
# Navigate to D1D system
sleep(1);
url = 'http://itpv2.transtron.fujitsu.com/';
driver.get(url);
sleep(1);

# %%
# Fill login information
sleep(1);
input_field = driver.find_element(By.ID, 'userid');
input_field.send_keys('izmb-free');

password_field = driver.find_element(By.ID, 'password');
password_field.send_keys('izumi001');

# %%
# Login
sleep(1);
btn_login = driver.find_element(By.ID, 'login');
btn_login.click();
sleep(1);

# %%
# Waiting the client loading
sleep(20);

# %%
# Get current time.
sleep(1);
currentTime = datetime.date.today();
currentDate = currentTime.day;
currentMonth = currentTime.month;
currentYear = currentTime.year;
scrappDataDate = currentDate - 1;

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# %%
# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
# Open Results
sleep(1);
results_menu = driver.find_element(By.ID, 'F40_dl');
results_menu.click();
sleep(2);

# %%
sleep(1);
vehicle_expense_report = driver.find_element(By.XPATH, '//*[@id="F45"]');
vehicle_expense_report.click();
sleep(5);

# %%
sleep(1);
checkbox = driver.find_element(By.XPATH, '//*[@id="F45_eigyousyoLevel_view"]/li/div/span/div/div');
checkbox.click();
sleep(2);

# %%
def getCurrentCalendarXPATH(calendar_id, year, month):
    return "//table[@id='" + calendar_id + "_calendar_start_" + year + '_' + month + "']//a";

# %%
sleep(1);
driver.find_element(By.ID, "F45_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F45_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F45', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
        
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
btn_dowload = driver.find_element(By.ID, 'F45_CSV_button');
btn_dowload.click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
driver_expense_report = driver.find_element(By.XPATH, '//*[@id="F46"]');
driver_expense_report.click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F46_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F46_from").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F46_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F46', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;

listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F46_csvout_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
driving_report_by_driver = driver.find_element(By.XPATH, '//*[@id="F42"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F42_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F42_from").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F42_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F42', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
        
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F42_csvout_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
driver_report_by_vehicle = driver.find_element(By.XPATH, '//*[@id="F43"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F43_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F43_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F43_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F43', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
        
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F43_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
yearly_mileage_report = driver.find_element(By.XPATH, '//*[@id="F44"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'F44_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F44_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);

# Open dropdown option Trip Start Year
trip_start_year = driver.find_element(By.ID, 'ComboBox_StartYear');
trip_start_year.click();

# Set the year to current year
select_trip_start_year = Select(trip_start_year);
select_trip_start_year.select_by_visible_text(str(currentYear));
sleep(1);

# Open dropdown option Trip Start Month
trip_start_month = driver.find_element(By.ID, 'ComboBox_StartMonth');
trip_start_month.click();

# Set the year to current year
select_trip_start_month = Select(trip_start_month);
select_trip_start_month.select_by_visible_text(str(currentMonth));
sleep(1);

# Open dropdown option Trip End Year
trip_end_year = driver.find_element(By.ID, 'ComboBox_EndYear');
trip_end_year.click();

# Set the year to current year
select_trip_end_year = Select(trip_end_year);
select_trip_end_year.select_by_visible_text(str(currentYear));
sleep(1);

# Open dropdown option Trip End Month
trip_end_month = driver.find_element(By.ID, 'ComboBox_EndMonth');
trip_end_month.click();

# Set the year to current year
select_trip_end_month = Select(trip_end_month);
select_trip_end_month.select_by_visible_text(str(currentMonth));
sleep(1);

# %%
sleep(1);
driver.find_element(By.ID, 'F33_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
digital_tachograph_data = driver.find_element(By.XPATH, '//*[@id="F49"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'F49_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '/html/body/article[1]/div/section/article/div[1]/article/div/div/div/ul/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F49_from").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F49_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F49', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F49_get_button').click();
sleep(2);

# %%
sleep(30);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
safety_driving_ranking_report = driver.find_element(By.XPATH, '//*[@id="F41"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'F41_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
# Check all branch
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F41_main_form"]/p[2]/div').click();

sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F41_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F41_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F41', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F41_CSV_button').click();
sleep(2);

# %%
sleep(30);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
driver_report = driver.find_element(By.XPATH, '//*[@id="F4B"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'F41_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F41_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F41_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F41', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F41_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
driving_report_by_vehicle = driver.find_element(By.XPATH, '//*[@id="F4C"]').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, 'F41_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F41_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F41_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F41', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F41_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
summary_table_by_department = driver.find_element(By.XPATH, '//*[@id="P21"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'P21_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="P21_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "P21_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="P21_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('P21', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'P21_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
# Open Menu
sleep(1);
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
sleep(1);
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
sleep(1);
safety_evaluation_results_table_by_vehicle = driver.find_element(By.XPATH, '//*[@id="F4G"]').click();
sleep(5);

# %%
sleep(1);
driver.find_element(By.ID, 'F4G_cancel_button').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F4G_eigyousyoLevel_view"]/li/div/span/div/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F4G_main_form"]/p[2]/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.XPATH, '//*[@id="F4G_main_form"]/p[3]/div').click();
sleep(2);

# %%
sleep(1);
driver.find_element(By.ID, "F4G_StartDate").click();

webDriverCurrentMonth = int(driver.find_element(By.XPATH, '//*[@id="F4G_calendar_start"]/div/div').text.split("/")[1]);

btn_previous = driver.find_element(By.ID, 'prev');
sleep(1);
btn_next = driver.find_element(By.ID, 'next');
sleep(1);

balancer = 0;
index = 0;

currentYearMonthCalendar = getCurrentCalendarXPATH('F4G', str(currentYear), str(currentMonth));

if webDriverCurrentMonth < currentMonth:
    balancer = currentMonth - webDriverCurrentMonth;
    
    while index < balancer:
        btn_next.click();
        index += 1;
else:
    balancer = webDriverCurrentMonth - currentMonth;
    while index < balancer:
        btn_previous.click();
        index += 1;
    
listDatesStart = driver.find_elements(By.XPATH, currentYearMonthCalendar);

for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;

# %%
sleep(1);
driver.find_element(By.ID, 'F4G_CSV_button').click();
sleep(2);

# %%
sleep(20);

# %%
win32api.MessageBox(0, 'Scrapping Completed.', 'Success', 0x00001000);
sleep(5);


