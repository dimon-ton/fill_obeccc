from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

df = pd.read_excel('fill_obeccc.xlsx')

print(len(df))

for i in range(len(df)+1):
    path = r'C:\Users\saich\Documents\basic_python\EP.13_automate_web\msedgedriver.exe'
    service = Service(path)
    driver = webdriver.Edge(service=service)
    url = 'https://app.contentcenter.obec.go.th/#/'
    driver.get(url)
    driver.set_window_size(1400,1000)

    time.sleep(2)

    elem = driver.find_element(By.XPATH, '/html/body/div/div/div[1]/div/div/div/div/nav/div/form/div/button')
    elem.click()

    # run loop email
    email = driver.find_element(By.ID, 'uemail')
    email.send_keys(df['e-mail'][i])

    password = driver.find_element(By.ID, 'upassword')
    cf_password = driver.find_element(By.ID, 'ucfpassword')

    password.send_keys('Chakpong1234')
    cf_password.send_keys('Chakpong1234')

    user_type = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[4]/div[1]/div/div/div/div[1]/div')

    dropdown = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[4]/div[1]/div/div/div/div[1]/input')
    user_type.click()

    dropdown.send_keys(Keys.ARROW_DOWN)
    dropdown.send_keys(Keys.ENTER)

    # loop prename
    pre = driver.find_element(By.ID, 'utitle')
    pre.send_keys(df.pre[i])

    # loop fname and lname
    fname = driver.find_element(By.ID, 'ufname').send_keys(df.fname[i])
    lname = driver.find_element(By.ID, 'ulname').send_keys(df.lname[i])
    #user_type.send_keys(Keys.ENTER)

    gender = driver.find_element(By.ID, "gender-male").click()


    addr = driver.find_element(By.ID, "address").send_keys('โรงเรียนวัดคลองชากพง สพป.ระยอง เขต 2')

    # Province
    prov = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div[7]/div[4]/div/div/div/div[1]")
    prov.click()

    prov_typing = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[7]/div[4]/div/div/div/div[1]/div[2]/div/input')

    prov_typing.send_keys('ระยอง')
    prov_typing.send_keys(Keys.ENTER)

    # District
    time.sleep(5)
    dist = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div[7]/div[5]/div/div/div/div[1]")
    dist.click()

    dist_typing = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[7]/div[5]/div/div/div/div[1]/div[2]/div/input')

    dist_typing.send_keys('แกลง')
    dist_typing.send_keys(Keys.ENTER)

    # Sub-District
    time.sleep(5)
    sub_dist = driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div[7]/div[6]/div/div/div/div[1]")
    sub_dist.click()

    sub_dist_typing = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[7]/div[6]/div/div/div/div[1]/div[2]/div/input')

    sub_dist_typing.send_keys('วังหว้า')
    sub_dist_typing.send_keys(Keys.ENTER)

    zip_code = driver.find_element(By.ID, 'zipcode')
    zip_code.click()
    zip_code.send_keys('21110')

    agree = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[2]/div[8]/div/div/div/input')
    agree.click()

    subscribe = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[3]/button[2]')
    subscribe.click()


    # quit programe
    driver.quit()