from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver import ActionChains


class GetReport:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.action = ActionChains(self.driver)

    def get_start_window(self):
        self.driver.get(
            "http://t2ru-tisdb-04/Reports/report/Standard%20Reports/Workflow/%D0%A0%D0%A0%D0%9B/"
            "WFL_%D0%B7%D0%B0%D0%B4%D0%B0%D0%BD%D0%B8%D1%8F%20%D0%BD%D0%B0%20SW%20%D1%80%D0%B0%D1"
            "%81%D1%88%D0%B8%D1%80%D0%B5%D0%BD%D0%B8%D1%8F%20%D0%A0%D0%A0%D0%9B")
        self.driver.maximize_window()
        time.sleep(22)

    def click_diskette(self):
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.TAB)
        self.action.send_keys(Keys.ENTER)
        self.action.send_keys(Keys.ARROW_DOWN)
        self.action.send_keys(Keys.ENTER)
        self.action.perform()
        time.sleep(12)

    def get_report_run(self):
        self.get_start_window()
        self.click_diskette()
        self.driver.quit()


if __name__ == "__main__":
    GetReport().get_report_run()
