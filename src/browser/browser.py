import time
from getpass import getpass
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from src.config import EASYJOB_SETTINGS, DEBUG, DATA_PATH, IHK_SETTINGS, AUSBILDUNGSABSCHNITT, BETREUEREMAIL
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from pandas import DataFrame, to_datetime
from src.utils import get_date_data, printProgressBar, cleanhtml
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
from sys import exit


class Browser:
    WEEKDAYMAPPING = [
        'MO',
        'DI',
        'MI',
        'DO',
        'FR',
        'SA',
        'SO'
    ]

    def __init__(self, driver):
        self.driver = driver
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    def check_exists(self, by=By.ID, value=None):
        try:
            self.driver.find_element(by, value)
        except NoSuchElementException:
            return False
        return True

    def find_element(self, by=By.ID, value=None, time=30, should_exit=True):
        try:
            elem = WebDriverWait(self.driver, time).until(
                EC.presence_of_element_located((by, value))
            )
            return elem
        except Exception as e:
            if should_exit:
                url = self.driver.current_url
                self.driver.quit()
                exit(f"Cant find field {value} on {url}" + str(e))

    def find_elements(self, by=By.ID, value=None):
        try:
            elem = WebDriverWait(self.driver, 30).until(
                EC.presence_of_all_elements_located((by, value))
            )
            return elem
        except Exception as e:
            url = self.driver.current_url
            self.driver.quit()
            exit(f"Cant find field {value} on {url}" + str(e))

    def get_dataEJ(self) -> dict:
        data_rows = self.find_elements(by=By.CLASS_NAME, value="eDVDataRow")

        if len(data_rows) < 5:
            exit("Not enough EasyJob Entry's")

        data = {
            'Datum': [],
            'Mitarbeiter': [],
            'Leistung': [],
            'Bezeichnung': [],
            'Anzahl': [],
            'Beschreibung': [],
        }
        print("Pulling Data")
        printProgressBar(0, len(data_rows), prefix='Progress:', suffix='Complete', length=50)
        for idx, row in enumerate(data_rows):
            row_id = row.get_attribute("id")
            row_num = re.search("\d{1,}", row_id).group(0)
            data['Datum'].append(
                self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C185").get_attribute('innerHTML')[:6] + '20' +
                self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C185").get_attribute('innerHTML')[6:])
            data['Mitarbeiter'].append(
                self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C250").get_attribute(
                    'innerHTML'))
            try:
                data['Leistung'].append(self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C251C3C106", time=1, should_exit=False).get_attribute('innerHTML').strip())
            except Exception:
                data['Leistung'].append(
                    self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C246C1C251").get_attribute('innerHTML').strip())

            data['Bezeichnung'].append(
                self.find_element(by=By.ID, value=f"eDVDDR{row_num}C252").get_attribute(
                    'innerHTML'))
            data['Anzahl'].append(
                self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C179").get_attribute('innerHTML'))
            description_element = self.find_element(by=By.ID, value=f"eDVCellDDR{row_num}C253C0C183")
            try:
                description = cleanhtml(description_element.find_element(by=By.TAG_NAME, value="div").get_attribute('innerHTML'))
            except Exception:
                description = ""
            data['Beschreibung'].append(description)

            printProgressBar(idx + 1, len(data_rows), prefix='Progress:', suffix='Complete', length=50)
        print("")
        return data

    def processEJ(self) -> DataFrame:
        print("Logging in to EJ")
        self.driver.get(EASYJOB_SETTINGS['url'])
        time.sleep(1)
        self.find_element(value='WinAuth').click()

        if not EASYJOB_SETTINGS['username']:
            EASYJOB_SETTINGS['username'] = input("username: ")
        if not EASYJOB_SETTINGS['password']:
            EASYJOB_SETTINGS['password'] = getpass("password: ")

        time.sleep(0.5)
        self.find_element(value='UserName').send_keys(EASYJOB_SETTINGS['username'])
        self.find_element(by=By.XPATH, value='//button[@type="next" and @x-tabindex="-1"]').click()

        for i in range(0, 3):
            self.find_element(value='UserName').send_keys(EASYJOB_SETTINGS['username'])
            self.find_element(value='Password').send_keys(EASYJOB_SETTINGS['password'])
            self.find_element(by=By.XPATH, value='//button[@type="submit" and @x-tabindex="-1"]').click()
            if self.driver.current_url != EASYJOB_SETTINGS['url']:
                break
            elif i == 2 and self.driver.current_url == EASYJOB_SETTINGS['url']:
                self.driver.quit()
                exit("password or username for Easyjob not valid.")
            else:
                print(f"password or username for Easyjob not valid. Try {i + 2}/3")
                EASYJOB_SETTINGS['username'] = input("username:")
                EASYJOB_SETTINGS['password'] = getpass("password: ")

        frame = self.find_element(by=By.XPATH, value='//iframe[@id="main"]')
        print("Search needed Data")
        self.driver.switch_to.frame(frame)
        self.find_element(by=By.XPATH,
                          value='//a[@id="menuToolbar_Menu_Stunden_menuStunden_ctl00_entryStdAbfrage_entryLb"]')
        self.driver.switch_to.default_content()
        self.driver.execute_script(
            "document.getElementById('main').contentWindow.document.getElementById("
            "'menuToolbar_Menu_Stunden_menuStunden_ctl00_entryStdAbfrage_entryLb').click();")
        original_window = self.driver.current_window_handle
        WebDriverWait(self.driver, 30).until(EC.number_of_windows_to_be(2))

        for window_handle in self.driver.window_handles:
            if window_handle != original_window:
                self.driver.switch_to.window(window_handle)
                break

        min_date = EASYJOB_SETTINGS['from_date']
        max_date = EASYJOB_SETTINGS['to_date']

        self.find_element(value='Master_selContent_dateRange_dateRange_dpDatumVon').clear()
        self.find_element(value='Master_selContent_dateRange_dateRange_dpDatumVon').send_keys(min_date)
        self.find_element(value='Master_selContent_dateRange_dateRange_dpDatumBis').clear()
        self.find_element(value='Master_selContent_dateRange_dateRange_dpDatumBis').send_keys(max_date)
        self.find_element(value='Master_selContent_btnFinde').click()

        if self.find_element(value='eDVDR0'):
            delta, min_date, max_date = get_date_data(min_date, max_date)
            for i in range(int(delta.days * 10 / 50)):
                data_rows = self.find_elements(by=By.CLASS_NAME, value="eDVDataRow")
                self.driver.execute_script('arguments[0].scrollIntoView();', data_rows[-1])
            data = self.get_dataEJ()
            self.driver.quit()
            df = DataFrame.from_dict(data)
            df["Datum"] = to_datetime(df["Datum"], format='%d.%m.%Y')
            df.sort_values(by='Datum', inplace=True, ascending=False)
            return df
        else:
            self.driver.quit()
            exit("Error")

    def getNeededIhkDates(self) -> datetime:
        self.processIhk()
        print("Getting Date")
        try:
            dateContainer = self.find_element(by=By.XPATH, value='//div[@class=" reihe"]', time=5, should_exit=False).\
                text
            datesString = dateContainer.split('\n')[1]
            lastDate = datesString.split(' - ')[1]
            lastDate = datetime.strptime(lastDate, '%d.%m.%Y') + timedelta(days=1)
        except Exception:
            datesString = self.find_elements(by=By.XPATH, value='//div[@class ="input-wrapper col-sm-12 col-md-8"]')[
                -1].text
            lastDate = datesString.split(' - ')[0]
            lastDate = datetime.strptime(lastDate, '%d.%m.%Y')

        self.driver.quit()
        return lastDate

    def createWeekEntrys(self, weeks: dict):
        self.processIhk()
        print("Creating IHK Entrys")

        printProgressBar(0, len(weeks), prefix='Creating IHK entry:', suffix='Complete', length=50)
        for idx, week in enumerate(weeks):
            self.find_element(by=By.XPATH, value='//form[@action="azubiHeftEditForm.jsp"]/button').click()
            ausbabschnitt = self.find_element(by=By.XPATH, value='//input[@name="ausbabschnitt" and @type="text"]')
            ausbabschnitt.clear()
            ausbabschnitt.send_keys(AUSBILDUNGSABSCHNITT)

            for i in ["", "2"]:
                ausbMail = self.find_element(by=By.XPATH, value=f'//input[@name="ausbMail{i}" and @type="text"]')
                ausbMail.clear()
                ausbMail.send_keys(BETREUEREMAIL)

            ausbinhalt = self.find_element(by=By.XPATH, value='//textarea[@name="ausbinhalt1"]')
            ausbinhalt.clear()
            ausbinhalt.send_keys(self.parseWeekDictToString(week))

            self.find_element(by=By.XPATH, value='//button[@name="save" and @type="submit"]').click()

            printProgressBar(idx + 1, len(weeks), prefix='Creating IHK entry:', suffix='Complete', length=50)

    def parseWeekDictToString(self, week: dict) -> str:
        ret = ""
        for day in week['days']:
            day_dt = datetime.strptime(day['date'], '%d.%m.%Y')
            ret += f"{self.WEEKDAYMAPPING[day_dt.weekday()]} "
            if int(day['time']['is'][0]) > 0:
                ret += f"[{day['time']['from']} bis {day['time']['to']} = {day['time']['is']}]"
            ret += ": "
            ret += ", ".join(day['descriptions'])
            ret += "\n"
        ret += " Gesamte Stunden: " + week['static']['whole_time']
        return ret

    def processIhk(self):
        print("Logging in to IHK")
        self.driver.get(IHK_SETTINGS['url'])

        if not IHK_SETTINGS['username']:
            IHK_SETTINGS['username'] = input("username:")
        if not IHK_SETTINGS['password']:
            IHK_SETTINGS['password'] = getpass("password: ")

        for i in range(0, 3):
            self.driver.get(IHK_SETTINGS['url'])
            self.find_element(by=By.NAME, value='login').send_keys(IHK_SETTINGS['username'])
            self.find_element(by=By.NAME, value='pass').send_keys(IHK_SETTINGS['password'])
            self.find_element(by=By.NAME, value='anmelden').click()
            bad_login = self.check_exists(by=By.XPATH, value='//div[@class="error"]')
            if not bad_login:
                break
            elif i == 2 and bad_login:
                self.driver.quit()
                exit("password or username for IHK not valid.")
            else:
                print(f"password or username for IHK not valid. Try {i + 2}/3")
                IHK_SETTINGS['username'] = input("username:")
                IHK_SETTINGS['password'] = getpass("password: ")

        self.find_element(by=By.XPATH, value='//a[@href="azubiHeft.jsp"]').click()


def setupBrowser():
    options = Options()
    options.headless = not DEBUG
    options.add_argument("--window-size=1920,1200")
    options.add_argument('--disable-blink-features=AutomationControlled')
    prefs = {"download.default_directory": DATA_PATH}
    options.add_experimental_option("excludeSwitches", ["enable-automation", 'enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return Browser(driver)
