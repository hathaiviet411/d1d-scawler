# %%
from selenium import webdriver;
from selenium.webdriver.common.by import By;
from selenium.webdriver.common.keys import Keys;
from selenium.webdriver.common.action_chains import ActionChains;
from webdriver_manager.chrome import ChromeDriverManager;
import datetime;
from time import sleep;

# %%
# Initialize Chrome
driver = webdriver.Chrome(ChromeDriverManager().install());
driver.maximize_window();

# %%
# Navigate to D1D system
url = 'http://itpv2.transtron.fujitsu.com/';
driver.get(url);
sleep(1);

# %%
# Fill login information
input_field = driver.find_element(By.ID, 'userid');
input_field.send_keys('izmb-free');

password_field = driver.find_element(By.ID, 'password');
password_field.send_keys('izumi001');

# %%
# Login
btn_login = driver.find_element(By.ID, 'login');
btn_login.click();
sleep(1);

# %%
# Waiting the client loading
sleep(20);

# %%
# Open Menu
btn_menu = driver.find_element(By.CLASS_NAME, 'b_btn-menu');
btn_menu.click();
sleep(2);

# %%
# Hover Function List
function_list = driver.find_element(By.ID, 'F100');
hover = ActionChains(driver).move_to_element(function_list);
hover.perform();

# %%
# Open Results
results_menu = driver.find_element(By.ID, 'F40_dl');
results_menu.click();
sleep(2);

# %%
# Get Safety Driving Ranking Data
safety_driving_ranking_data = driver.find_element(By.XPATH, '//*[@id="F41"]');
safety_driving_ranking_data.click();
sleep(2);

# %%
# Check all branch
select_all_checkbox = driver.find_element(By.XPATH, '//*[@id="F41_eigyousyoLevel_view"]/li/div/span/div/div');
select_all_checkbox.click();

# %%
# Check total safety driving ranking report
total_safety_driving_ranking_report = driver.find_element(By.XPATH, '//*[@id="F41_main_form"]/p[2]/div');
total_safety_driving_ranking_report.click();

# %%
# Get current time.
currentTime = datetime.date.today();
currentDate = currentTime.day;
scrappDataDate = currentDate - 1;

# %%
# Set start day
start_day = driver.find_element(By.ID, "F41_StartDate").click();
listDatesStart = driver.find_elements(By.XPATH, "//table[@id='F41_calendar_start_2021_10']//a");
for date in listDatesStart:
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;
    else:
        print('ERROR');

# %%
# Set end day
end_day = driver.find_element(By.ID, "F41_EndDate").click();
listDatesEnd = driver.find_elements(By.XPATH, "//table[@id='F41_calendar_end_2021_10']//a");
for date in listDatesEnd:
    print(date.text);
    dateName = date.text;
    if dateName == str(scrappDataDate):
        date.click();
        break;
    else:
        print('ERROR');

# %%
# Output data.
btn_output = driver.find_element(By.ID, 'F41_CSV_button');
btn_output.click();


