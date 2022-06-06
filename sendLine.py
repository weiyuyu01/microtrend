
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import win32con,win32gui
import pymysql
import os
import pyimgur
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import time 
import sys


def db_init():
    db = pymysql.connect(
        host='127.0.0.1',
        user='root',    
        password='root',
        port=3306,
        db='test'
    )
    cursor = db.cursor(pymysql.cursors.DictCursor)
    return db, cursor

def drop_table():
    db, cursor = db_init()
    sql = """
        drop table `friends`
    """
    cursor.execute(sql)
    db.commit()
    db.close()

def create_table():
    db, cursor = db_init()
    sql = """
    CREATE TABLE IF NOT EXISTS friends (
    ID INT AUTO_INCREMENT PRIMARY KEY,
    name varchar(100) collate utf8mb4_general_ci,
    datamid varchar(100),
    icon BLOB NOT NULL );
    """
    cursor.execute(sql)
    db.commit()
    db.close()
    
       
#study_1
def driver_setting(login_method):
    chrome_options = Options()
    goal = os.path.abspath(r'C:\\ROline\\extension_2_4_9_0.crx')
    chrome_options.add_extension(goal)
    # chrome_options.add_extension(r'(C:\\ROline\\2.5.1_0.crx')
    driver = webdriver.Chrome(ChromeDriverManager().install() ,chrome_options=chrome_options)
    driver.get("chrome-extension://ophjlpahpchlmihnnnihgmmeilfjmjjc/index.html")
    if login_method == 'login':
        mail_login = WebDriverWait(driver, 20, 0.5).until(
            EC.presence_of_element_located((By.XPATH , '//li[@class = "TabsEmail"]'))
            )
        ActionChains(driver).click(mail_login).perform()
    else:
        qrcode_login = WebDriverWait(driver, 20, 0.5).until(
            EC.presence_of_element_located((By.XPATH , '//li[@class = "TabsQR"]'))
            )
        ActionChains(driver).click(qrcode_login).perform()
    return driver

#study_2
def line_login(email , password):
    
    ### 登入line ###
    WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="line_login_email"]'))
        )
    email_box = driver.find_element(By.XPATH, '//*[@id="line_login_email"]')
    email_box.send_keys(email)
    pwd_box=driver.find_element(By.XPATH, '//*[@id="line_login_pwd"]')
    ActionChains(driver).click(pwd_box).send_keys(password).perform()
    time.sleep(2)
    pwd_box.send_keys(Keys.ENTER)


#study_3
def get_friends_list():
    
    ### 等待賴畫面出現，並點擊好友圖示 ###
    WebDriverWait(driver, 1000, 0.5).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="leftSide"]/div[10]/ul/li[2]'))
        ).click()
    
    ### 滾動找出所有好友 ###
    target = WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="contact_wrap_friends"]/div/h2'))
        )
    ActionChains(driver).click(target).send_keys(Keys.END).perform()
    
    time.sleep(3)
    
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    #contact_wrap_friends
    list_soup = soup.find('div', {'id':'contact_wrap_friends'}).find('ul',{'class':'MdCMN03List ExHover'})
    goal_soup=list_soup.find_all('li',{'class':'mdMN02Li'})
    name_list = []
    for name in goal_soup: name_list.append(name.get('title'))
    id_list = []
    for id in goal_soup: id_list.append(id.get('data-mid'))
    # icon_list = []
    # icon_soup=list_soup.find_all('img',{'class':'_profile_img'})
    # for icon in icon_soup : icon_list.append(icon.get('src'))
    data_list = zip(name_list,id_list) #icon_list
    return data_list

def add_friends_to_DB(list):
    db, cursor = db_init()
    sql = """
        INSERT INTO `friends` (`name` , `datamid` , `icon`) VALUES (%s , %s, %s)
    """
    cursor.executemany(sql,list)
    db.commit()
    db.close()
    
def send_message(uid ,  message):
    
    ### 等待賴畫面出現，並點擊好友圖示 ###
    WebDriverWait(driver, 1000, 0.5).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="leftSide"]/div[10]/ul/li[2]'))
        ).click()
    ### 滾動找出所有好友 ###
    target = WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="contact_wrap_friends"]/div/h2'))
        )
    ActionChains(driver).click(target).send_keys(Keys.END).perform()
    
    ### 定位出要送出訊息的好友(by id) ###
    xpath = "//li[@data-mid='{}']".format(uid)
    friend = WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH,xpath))
        )
    ActionChains(driver).click(friend).perform()
    #賣樂圈:u39dda45e1eb2c29eb765ba3240032b78  
    
    if '.' in message :
        CLIENT_ID = "d596e2b1bac7520"
        image = 'anya.jpg'
        PATH = (f"C:\\ROline\\{image}") 
        im = pyimgur.Imgur(CLIENT_ID)
        upload_url = im.upload_image(PATH)
        image_url=upload_url.link
        intput_value = driver.find_element(By.XPATH , '//*[@id="_chat_room_input"]')
        js = f"""
        var i = document.createElement('img');
        i.setAttribute("src", "{image_url}");
        i.setAttribute("style" , "max-width:250px;max-height:75px;");
        var target_ele = arguments[0];
        target_ele.insertAdjacentElement("beforeend" , i );
        """
        driver.execute_script(js , intput_value)
        target = driver.find_element(By.XPATH,'//*[@id="_chat_room_input"]')
        ActionChains(driver).click(target).send_keys(Keys.ENTER).perform()   
        time.sleep(5)
        print(' SUCCESS: FINISH: 發送完成') 
        driver.quit()
    else:    
        ### 送出訊息 ###
        text = WebDriverWait(driver, 20, 0.5).until(
            EC.presence_of_element_located((By.XPATH , '//*[@id = "_chat_room_input"]'))
            )
        ActionChains(driver).click(text).send_keys(message).send_keys(Keys.ENTER).perform()
        time.sleep(5)
        print(' SUCCESS: FINISH: 發送完成')
        driver.quit()
        
    
    
### 利用套件接管windows畫面 ###
def send_img():
    WebDriverWait(driver, 1000, 0.5).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="leftSide"]/div[10]/ul/li[2]'))
        ).click()
    
    ### 滾動找出所有好友 ###
    target = WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="contact_wrap_friends"]/div/h2'))
        )
    ActionChains(driver).click(target).send_keys(Keys.END).perform()
    
    ### 定位出要送出訊息的好友(by id) ###
    WebDriverWait(driver, 30, 0.5).until(
        EC.element_to_be_clickable((By.XPATH,'//li[@data-mid="u2e46caaadff28b97a6183de5eaa6d3cc"]'))
        ).click()
    WebDriverWait(driver, 20, 1).until(
        EC.presence_of_element_located((By.XPATH , '//*[@id="_chat_room_plus_btn"]'))
        ).click()
    time.sleep(2)
    
    dialog = win32gui.FindWindow("#32770", "開啟")
    ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, "ComboBoxEx32", None)  
    ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, "ComboBox", None)  
    edit = win32gui.FindWindowEx(ComboBox, 0, "Edit", None) 
    file_path = r'F:\ROline樂賴\777.png'
    
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, None, file_path)        
    win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1,"{ENTER}") 
 
    
def send_img_js(image):
    
    ### 產生圖片url ###
    goal = os.path.abspath(image)
    driver2 = webdriver.Chrome(ChromeDriverManager().install())
    driver2.get("https://imgur.com/")
    driver2.implicitly_wait(2)
    WebDriverWait(driver2, 1000, 0.5).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[1]/div/div[1]/div[2]/div[3]/a[1]'))
        ).click()
    driver2.find_element(By.XPATH , '//*[@id="file-input"]').send_keys(goal) 
    WebDriverWait(driver2 , 100, 0.5).until(
        EC.presence_of_element_located((By.XPATH,'//div[@class="PostContent-imageWrapper-rounded"]'))
        )
    time.sleep(3)
    html = driver2.page_source
    soup = BeautifulSoup(html, 'html.parser')
    block =soup.find('div',{'class':'PostContent-imageWrapper'})
    image_url = block.find('div',{'class':'PostContent-imageWrapper-rounded'}).find('img').get('src')
    driver2.quit()
    WebDriverWait(driver, 1000, 0.5).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="leftSide"]/div[10]/ul/li[2]'))
        ).click()
    
    ### 滾動找出所有好友 ###
    target = WebDriverWait(driver, 30, 0.5).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="contact_wrap_friends"]/div/h2'))
        )
    ActionChains(driver).click(target).send_keys(Keys.END).perform()
    
    ### 定位出要送出訊息的好友(by id) ###
    WebDriverWait(driver, 30, 0.5).until(
        EC.element_to_be_clickable((By.XPATH,'//li[@data-mid="u39dda45e1eb2c29eb765ba3240032b78"]'))
        ).click()
    
    ### 輸入圖片並送出 ###
    intput_value = driver.find_element(By.XPATH , '//*[@id="_chat_room_input"]')
    js = f"""
    var i = document.createElement('img');
    i.setAttribute("src", "{image_url}");
    i.setAttribute("style" , "max-width:250px;max-height:75px;");
    var target_ele = arguments[0];
    target_ele.insertAdjacentElement("beforeend" , i );
    """
    driver.execute_script(js , intput_value)
    target = driver.find_element(By.XPATH,'//*[@id="_chat_room_input"]')
    ActionChains(driver).click(target).send_keys(Keys.ENTER).perform()
    


if __name__ =='__main__':
    
    ### 系統指令輸入
    
    driver = driver_setting(sys.argv[1])
    line_login(sys.argv[2] , sys.argv[3])
    send_message(sys.argv[4],sys.argv[5])
    
    # add_friends_to_DB(list)
    # send_img()
    # list = get_friends_list()
    
    # drop_table() 
    # create_table()
    
    