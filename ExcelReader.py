# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import Util

class ExcelReader(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(30)
        self.base_url = "http://192.168.0.22:8801/opt/a"
        self.verificationErrors = []
        self.accept_next_alert = True

    def login(self):
        driver = self.driver
        driver.get(self.base_url + "/")
        driver.find_element_by_id("username").clear()
        driver.find_element_by_id("password").clear()
        time.sleep(4)
        driver.find_element_by_id("username").send_keys("baomi")
        driver.find_element_by_id("password").send_keys("123456")
        #driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()

        time.sleep(1)



        #切换窗体句柄

        time.sleep(60)


    def is_element_present(self, how, what):
        try:
            self.driver.find_element(by=how, value=what)
        except NoSuchElementException as e:
            return False
        return True

    def is_alert_present(self):
        try:
            self.driver.switch_to_alert()
        except NoAlertPresentException as e:
            return False
        return True

    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally:
            self.accept_next_alert = True

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    unittest.main()









