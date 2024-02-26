from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

web = webdriver.Chrome()
web.get("https://wa.me/972505978097")
web.implicitly_wait(3)
#web.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/div/div/div/div[1]/div/div[1]/a').click()
sleep(100)
web.close()