
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
import time

# 打开浏览器
driver = webdriver.Chrome()

# 最大化浏览器窗口
driver.maximize_window()

# 广东省人民医院智慧病理平台登录界面(测试环境)
driver.get("http://192.168.0.99/pypacs/#/login")

WebDriverWait(driver, 10).until(lambda x: x.find_element_by_xpath('//*[@id="app"]/div[1]/div/div/form/div[1]/div/div[1]/input'))

# 输入用户名和密码
username = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div/form/div[1]/div/div[1]/input').send_keys("super")
password = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div/form/div[2]/div/div[1]/input').send_keys("admin")

# 点击登录按钮
login_button = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div/div/form/button').click()

# 点击关闭弹窗继续登录
Continue_Login = driver.find_element_by_xpath('/html/body/div[2]/div/div[1]/button/i').click()

# 等待页面加载完成
time.sleep(5)

# 关闭浏览器
driver.quit()