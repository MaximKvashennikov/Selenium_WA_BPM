from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver import ActionChains
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WaBPM:
    def __init__(self, responsible, executor, region_name, mr_name, start_date_for_bpm, end_date_for_bpm,
                 start_time_for_bpm,
                 end_time_for_bpm, rrl_list_sw_file, influence_list_sw_file, path_to_driver=None):
        self.driver = webdriver.Chrome(executable_path=path_to_driver)
        self.action = ActionChains(self.driver)
        self.responsible = responsible
        self.executor = executor
        self.region_name = region_name
        self.mr_name = mr_name
        self.start_date_for_bpm = start_date_for_bpm
        self.end_date_for_bpm = end_date_for_bpm
        self.start_time_for_bpm = start_time_for_bpm
        self.end_time_for_bpm = end_time_for_bpm
        self.rrl_list_sw_file = rrl_list_sw_file
        self.influence_list_sw_file = influence_list_sw_file

    def get_start_window(self):
        self.driver.get("https://bpm.tele2.ru/0/Nui/ViewModule.aspx#SectionModuleV2/UsrWorkOrderSection")
        self.driver.maximize_window()

    def add_work(self):
        try:
            wait = WebDriverWait(self.driver, 20)
            element = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="UsrWorkOrderSectionSeparateModeAddRecordButtonButton-textEl"]')))
            print(element)
            self.driver.find_element_by_xpath(
                '//*[@id="UsrWorkOrderSectionSeparateModeAddRecordButtonButton-textEl"]').click()
            self.driver.implicitly_wait(10)
        except Exception:
            self.driver.find_element_by_id('UsrWorkOrderSectionSeparateModeAddRecordButtonButton-textEl').click()
            self.driver.implicitly_wait(10)

    def get_sr(self):
        """ Получение номера заявки """

        time.sleep(3)
        text_wa = self.driver.find_element_by_xpath('//*[@id="MainHeaderSchemaPageHeaderCaptionLabel"]').text
        time.sleep(2)
        if text_wa not in "WA":
            time.sleep(4)
            text_wa = self.driver.find_element_by_xpath('//*[@id="MainHeaderSchemaPageHeaderCaptionLabel"]').text
        time.sleep(2)

        return text_wa

    def input_config(self):
        time.sleep(4)
        wait = WebDriverWait(self.driver, 30)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="UsrWorkOrderPageUsrShortDescriptionComboBoxEdit-el"]')))

        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrShortDescriptionComboBoxEdit-el"]').send_keys(
            "Изменение настроек/конфигурации")
        time.sleep(3)
        self.action.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        self.action.send_keys(Keys.ENTER)
        time.sleep(1)
        self.action.perform()

    def input_access(self):
        self.driver.execute_script("window.scrollTo(0, 0)")
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrWODestinationLookupEdit-el"]').send_keys(
            "Сети доступа", Keys.ENTER)
        time.sleep(5)

        self.driver.find_element_by_xpath('//*[@id="searchEdit-el"]').click()
        time.sleep(1)

        """ срез [:-2] позволяет избежать ошибки stale element reference: element is not 
        attached to the page document, до обногвления BPM ее не было"""

        access_list = [item.text for item in self.driver.find_element_by_xpath(
            '//*[@id="grid-grid-wrap"]').find_elements_by_tag_name("div")[:-2]]
        access_list = [web_access.split("\n")[0] for web_access in access_list if
                       "\n" in web_access and "Сети доступа" in web_access]
        print(access_list)

        def search_index():
            for index, item in enumerate(access_list):
                if item == "Сети доступа":
                    return index

        access_position = int(search_index())
        print(access_position)
        xpath_str = '//*[@id="grid-grid-wrap"]/div[{}]'.format(access_position + 2)

        self.driver.find_element_by_xpath(xpath_str).click()
        # self.driver.find_element_by_xpath('//*[@id="grid-item-0418133e-9445-4897-8f72-7d60011a7836"]/div[1]').click()
        self.driver.implicitly_wait(5)
        time.sleep(1)

        try:
            self.driver.execute_script(
                "document.getElementById('selectionControlsContainerLookupPage').children[0].click()")
            time.sleep(2)
            # self.driver.find_element_by_xpath('//*[@id="selectionControlsContainerLookupPage"]/div[0]').click()
            # self.driver.implicitly_wait(10)
        except Exception:
            self.driver.find_element_by_id('grid-item-0418133e-9445-4897-8f72-7d60011a7836').click()
            self.driver.implicitly_wait(10)
        time.sleep(1)

    def job_description(self):
        description_str = "Увеличение ёмкости РРЛ {}:\n{}".format(str(self.mr_name), self.rrl_list_sw_file)
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrDescriptionMemoEdit-el"]').send_keys(
            description_str)

    def input_executor(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrOwnerLookupEdit-el"]'). \
            send_keys(self.responsible, Keys.ENTER)
        time.sleep(2)
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[2]').click()
        self.driver.implicitly_wait(10)
        time.sleep(1)
        self.driver.execute_script(
            "document.getElementById('selectionControlsContainerLookupPage').children[0].click()")
        time.sleep(2)

    def click_pop_up_window(self):
        try:
            time.sleep(1)
            self.driver.find_element_by_xpath('//*[@id="t-comp0-wrap"]/span[1]').click()
            time.sleep(1)
        except Exception:
            time.sleep(1)

    def input_time(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeStartDateEdit-el"]'). \
            send_keys(self.start_date_for_bpm, Keys.TAB)
        time.sleep(1)

        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeStartTimeEdit-el"]'). \
            send_keys(self.start_time_for_bpm, Keys.TAB)
        time.sleep(1)

        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeEndDateEdit-el"]'). \
            send_keys(self.end_date_for_bpm, Keys.TAB)
        self.action.send_keys(Keys.TAB).perform()
        time.sleep(2)
        self.click_pop_up_window()

        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeEndTimeEdit-el"]').click()
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeEndTimeEdit-el"]').send_keys(
            4 * Keys.BACKSPACE)

        time.sleep(1)
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrPlannedTimeEndTimeEdit-el"]'). \
            send_keys(self.end_time_for_bpm, Keys.TAB)

        time.sleep(3)

    def input_influence_time(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrExpectTimeTimeEdit-el"]'). \
            send_keys("1:00", Keys.TAB)
        time.sleep(1)

    def geography_of_influence(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageIsAffectOnSubscriberCheckBoxEdit-el"]').click()
        time.sleep(1)
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageGeoInfluenceComboBoxEdit-el"]'). \
            send_keys("Часть региона")
        self.action.send_keys(Keys.ENTER)
        time.sleep(1)
        self.action.send_keys(Keys.ARROW_DOWN).perform()
        time.sleep(1)

    def input_plan_fields(self):

        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrRisksServiceMemoEdit-el"]').send_keys(
            self.influence_list_sw_file)
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrWorkPlanMemoEdit-el"]'). \
            send_keys(
            "1. Проверка параметров: BBE, ES, SES, UAS, аварийных сообщений на затрагиваемых модулях корзины, "
            "приемных уровней на пролете.\n"
            "2. Изменение конфигурации терминала FarEnd.\n"
            "3. Изменение конфигурации терминала NearEnd.\n"
            "4. Проверка параметров: BBE, ES, SES, UAS, аварийных сообщений на затрагиваемых модулях корзины, "
            "приемных уровней на пролете.\n"
            "5. Проверка доступности БС. В случае недоступности БС - возврат в предыдущим настройкам, оповещение NOC. "
            "Проверка аварий и основных KPI.")
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageUsrRollbackPlanMemoEdit-el"]'). \
            send_keys("1. При слете конфигурации корзины - возврат к первоначальным настройкам.\n"
                      "2. При падении пролета - откат к первоначальной конфигурации РРЛ.\n"
                      "3. При отсутствии возможности отката - оповещение сотрудников отдела эксплуатации"
                      " регионального филиала, оповещение сотрудников NOC, выезд бригады на сайт.\n"
                      "4. При потере управляющего линка с корзиной РРЛ - оповещение сотрудников отдела "
                      "эксплуатации регионального филиала, оповещение сотрудников NOC, выезд бригады на сайт.\n")
        time.sleep(1)
        self.action.send_keys(Keys.TAB).perform()

        '''Скрол до конца страницы'''
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    def click_input_executor(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWOExecutorDetailAddRecordButtonButton-imageEl"]').click()
        self.driver.find_element_by_xpath('//*[@id="UsrContactLookupEdit-el"]'). \
            send_keys(self.executor, Keys.ENTER)
        time.sleep(1)
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[2]').click()

        self.driver.implicitly_wait(10)
        self.driver.execute_script(
            "document.getElementById('selectionControlsContainerLookupPage').children[0].click()")
        time.sleep(2)
        self.driver.find_element_by_xpath(
            '//*[@id="grid-UsrWOExecutorDetailDataGridGrid-wrap"]/div[2]/div[2]/span[1]/span').click()
        time.sleep(2)

    def input_region(self, region_name):
        self.driver.find_element_by_xpath('//*[@id="UsrWORegionDetailV2AddRecordButtonButton-imageEl"]').click()
        time.sleep(2)
        self.driver.find_element_by_xpath('//*[@id="searchEdit-el"]'). \
            send_keys(region_name, Keys.ENTER)
        time.sleep(2)
        self.action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        city_list = [item.text for item in self.driver.find_element_by_xpath(
            '//*[@id="grid-grid-wrap"]').find_elements_by_tag_name("div")[:-2]]

        if region_name not in ("Омск", "Томск"):
            city_list = [city.split("\n")[0] for city in city_list if "\n" in city and region_name in city]

        else:
            city_list = [city.split("\n")[0] for city in city_list if "\n" in city and "мск" in city]

        print(city_list)

        def search_index():
            for index, item in enumerate(city_list):
                if item == region_name:
                    return index

        region_position = int(search_index())
        print(region_position)
        xpath_str = '//*[@id="grid-grid-wrap"]/div[{}]/div[1]/span/input'.format(region_position + 2)

        self.driver.find_element_by_xpath(xpath_str).click()
        self.driver.execute_script(
            "document.getElementById('selectionControlsContainerLookupPage').children[0].click()")
        time.sleep(1)

        '''Скрол до конца страницы'''
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    def click_on_influence(self):
        self.driver.find_element_by_xpath(
            '//*[@id="UsrCustomerServiceInWODetailAddRecordButtonButton-imageEl"]').click()
        time.sleep(1)

        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[3]/div[1]/span/input').click()
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[4]/div[1]/span/input').click()
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[5]/div[1]/span/input').click()
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[10]/div[1]/span/input').click()
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[11]/div[1]/span/input').click()
        self.driver.find_element_by_xpath('//*[@id="grid-grid-wrap"]/div[12]/div[1]/span/input').click()

        self.driver.execute_script(
            "document.getElementById('selectionControlsContainerLookupPage').children[0].click()")
        time.sleep(2)

    def save_work(self):
        self.driver.find_element_by_xpath('//*[@id="UsrWorkOrderPageSaveButtonButton-textEl"]').click()
        try:
            time.sleep(3)
            self.driver.find_element_by_xpath('//*[@id="t-comp0-wrap"]/span[1]').click()
            time.sleep(3)
        except Exception:
            print("Нет пересечений")

    def authorization_bpm(self):
        time.sleep(3)
        action = ActionChains(self.driver)
        action.key_down(Keys.ALT).key_down(Keys.ENTER).perform()
        time.sleep(10)

    def run_wa(self):
        self.get_start_window()
        self.authorization_bpm()
        self.add_work()
        # text_wa = self.get_sr()
        self.input_config()
        text_wa = self.get_sr()
        self.input_access()
        self.job_description()
        self.input_executor()
        self.input_time()
        self.click_pop_up_window()
        self.input_influence_time()
        self.geography_of_influence()
        self.input_plan_fields()
        self.click_input_executor()
        self.input_region(region_name=self.region_name)
        self.click_on_influence()
        self.save_work()
        self.driver.quit()

        return text_wa
