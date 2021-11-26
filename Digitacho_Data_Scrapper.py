import requests;
from selenium import webdriver;
from selenium.webdriver.common.by import By;
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep;
from bs4 import BeautifulSoup;
import os;
import zipfile;

# Initialize Chrome
driver = webdriver.Chrome();

# Navigate to D1D system
url = 'http://itpv2.transtron.fujitsu.com/';
driver.get(url);
sleep(1);

# Fill login information
input_field = driver.find_element(By.ID, 'userid');
input_field.send_keys('izmb-free');

password_field = driver.find_element(By.ID, 'password');
password_field.send_keys('izumi001');

# Login
btn_login = driver.find_element(By.ID, 'login');
btn_login.click();
sleep(1);

# Waiting the client loading
sleep(20);

# Open Menu
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# Hover Function List
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# Open Results
results_menu = driver.find_element(By.ID, 'F40_dl');
results_menu.click();

# Get Digital Tachograph Data
digital_tachograph_data = driver.find_element(By.XPATH, '//*[@id="F40_ul"]/li[7]');
digital_tachograph_data.click();

# Check all branch
checkbox_select_all_branch = driver.find_element(By.XPATH, '//*[@id="F49_eigyousyoLevel_view"]/li/div/span/div/div');
checkbox_select_all_branch.click();
sleep(2);

# Adjust trip start date
trip_start_day = driver.find_element(By.ID, 'F49_from');
trip_start_day.click();

calendar_field = driver.find_element(By.XPATH, '//*[@id="F49_calendar_start_2021_11"]/tbody/tr[4]/td[3]');
ActionChains(driver).move_to_element(calendar_field).click().perform();

# Dowload
btn_get_data = driver.find_element(By.ID, 'F49_get_button');
btn_get_data.click();