import json
import traceback
import sys
import pandas as pd
from pathlib import Path
import csv
import time
import requests
import os
from typing import Callable, Any, Union, Iterable
import parsedatetime
import datetime
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class AuthenticationError(ValueError):
    pass

class WebDriver:
    # most credits go here
    # https://jpmelos.com/articles/how-use-chrome-selenium-inside-docker-container-running-python/
    def __init__(self, headless=True, download_dir='/tmp/xero_custom_reports'):
        try:
            print("making tmp directory for saving excels", download_dir)
            os.makedirs(download_dir)
        except FileExistsError:
            pass

        self.download_dir = download_dir
        self.options = webdriver.ChromeOptions()

        self.options.add_argument('--disable-extensions')

        if headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--disable-gpu')
            self.options.add_argument('--no-sandbox')

        # to make the button clickable
        self.options.add_argument('--window-size=1920,1080')
        user_agent = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_5)"
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36")
        self.options.add_argument('--user-agent={}'.format(user_agent))

        self.options.add_experimental_option(
            'prefs', {
                'download.default_directory': self.download_dir,
                'download.prompt_for_download': False,
                'download.directory_upgrade': True,
                'intl.accept_languages': 'en,en_US',
                'safebrowsing.enabled': True
            }
        )

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, *args, **kwargs):
        self.close()

    def open(self):
        self.driver = webdriver.Chrome(chrome_options=self.options)
        self.driver.implicitly_wait(10)

    def close(self):
        self.driver.quit()

    def enable_download_in_headless_chrome(self):
        # downloading files in headless mode doesn't work!
        # file isn't downloaded and no error is thrown
        #add missing support for chrome "send_command"  to selenium webdriver
        # https://bugs.chromium.org/p/chromium/issues/detail?id=696481#c39
        self.driver.command_executor._commands["send_command"] = ("POST", '/session/{}/chromium/send_command'.format(self.driver.session_id))
        params = {
            'cmd': 'Page.setDownloadBehavior',
            'params': {'behavior': 'allow', 'downloadPath': self.download_dir}
        }
        self.driver.execute("send_command", params)

    def login(self, username, password, url="https://login.xero.com"):
        self.driver.get(url)

        email_field = self.driver.find_element_by_id("email")
        email_field.clear()
        email_field.send_keys(username)

        pass_field = self.driver.find_element_by_id("password")
        pass_field.clear()
        pass_field.send_keys(password)
        pass_field.send_keys(Keys.RETURN)
        time.sleep(2)

        title = self.driver.title
        if "Xero | Dashboard" not in title:
            raise AuthenticationError(
                ("Probably didn't authenticate sucesfully. "
                 "The title says '{}'. Check your credentials and account_id!").format(title))

    def list_reports(self, account_id):
        self.driver.get(
            "https://reporting.xero.com/{}/v2/ReportList/CustomReports?".format(
            account_id))
        # it's a json returned in html
        js = json.loads(self.driver.find_element_by_tag_name('pre').text)
        return {r["id"]: r["name"] for r in js['reports']}
        # print report ids and names here?

    @staticmethod
    def account_id_from_url(url):
        try:
            return re.search(r"(?!=.com/)!\w+", url).group()
        except AttributeError:
            raise ValueError("Couldn't find account_id in {}".format(url))

    def download_report(
            self,
            report_id,
            account_id,
            from_date=None,
            to_date=None,
            delay_seconds=15):
        """
        """
    
        # the first string after xero.com/ is the account id
        report_download_template = (
            "https://reporting.xero.com/"
            "{account_id}/v1/Run/"
            "{report_id}?isCustom=True").format(account_id=account_id, report_id=report_id)

        print("getting report from ", report_download_template,
              "from_date", from_date,
              "to_date", to_date)
        self.driver.get(report_download_template)
        self.enable_download_in_headless_chrome()
        print("Waiting for page to load")
        time.sleep(3)
        self.update_date_range(from_date, to_date, delay_seconds)

        # click export btn so that excel btn is rendered
        export_btn = self._locate_export_button()
        export_btn.click()
        time.sleep(1)

        excel_btn = self._locate_export_to_excel_button()
        excel_btn.click()
        # the sleeps are experimental and it might happen that the file won't be downloaded
        time.sleep(3)
        while True:
            report_path = glob_excels(self.download_dir)
            if not report_path:
                print("Waiting for the report to be downloaded")
                time.sleep(3)
            else:
                print("Report downloaded to ", report_path)
                time.sleep(3)
                break
                    
                    
    def direct_url(self,
                   account_id,
                   url = None):
        
        if url == None:
            raise ValueError("No URL provided.")
        else:
            explicit_company_url = (
                    "https://reporting.xero.com/"
                    "{account_id}/summary").format(account_id=account_id)
            print("Localise company ", explicit_company_url)
            self.driver.get(explicit_company_url)
            print("Waiting for page to load")
            time.sleep(3)
            print("Getting report from ", url,
                  " for company: ", account_id)
            self.driver.get(url)
            print("Waiting for page to load")
            time.sleep(3)
            self.enable_download_in_headless_chrome()
            self.driver.get(url.replace("Report.aspx?", "ExcelReport.aspx?", 1))
            time.sleep(3)
            count = 0 #for debug
            while True:
                report_path = glob_excels(self.download_dir)
                if count > 10: #for debug
                    raise ValueError("Looping too many times") #for debug
                elif not report_path:
                    print("Waiting for the report to be downloaded")
                    time.sleep(3)
                    count += 1
                else:
                    print("Report downloaded to ", report_path)
                    time.sleep(3)
                    break

    def _locate_export_button(self):
        print("Looking for export button")
        for btn in  self.driver.find_elements_by_tag_name('button'):
            if btn.get_attribute('data-automationid') == 'report-toolbar-export-button':
                print("Found")
                return btn
        raise KeyError("Couldn't find export menu. "
                       "The underlying html/css probably changed and the code needs to be adjusted")
    def _locate_export_to_excel_button(self):
        print("Looking for excel button")
        for btn in self.driver.find_elements_by_tag_name('button'):
            if btn.get_attribute('data-automationid') == 'report-toolbar-export-excel-menuitem--body':
                print("Found")
                return btn

        raise KeyError("Couldn't find export button. "
                       "The underlying html/css probably changed and the code needs to be adjusted")

    def update_date_range(self,
                          from_time: Union[str, None],
                          until_time: Union[str, None],
                          delay_seconds: int=15):
        # update From field
        if from_time:
            print("Updating from time")
            from_input_field = self.driver.find_element_by_id("dateFieldFrom-inputEl")
            print("Found input field", from_input_field)
            time.sleep(2.5)
            from_input_field.send_keys(
                Keys.BACKSPACE * len(from_input_field.get_attribute("value")))
            time.sleep(2.5)

            from_input_field.send_keys(from_time)
            # press the update button
            print("Field updated")

        if until_time:
            # update To field
            print("Updating until time")
            until_input_field = self.driver.find_element_by_id("dateFieldTo-inputEl")
            time.sleep(2.5)
            # clear the input field
            until_input_field.send_keys(
                Keys.BACKSPACE * len(until_input_field.get_attribute("value")))

            time.sleep(2.5)
            # input the datetime
            until_input_field.send_keys(until_time)
            print("field updated")
            time.sleep(2.5)
            # take the first div that satisfies the condition
            print("Looking for update button")
            update_btn = next(filter(
                lambda div: div.get_attribute("data-automationid") == "date-toolbar-update-button",
                self.driver.find_elements_by_tag_name("div")))
            print("Found", update_btn)
            print("Clicking")
            update_btn.click()
            # wait 15 seconds to the report is, hopefully, updated we can't really
            # use explicit waits baked into selenium because the buttons do not
            # have unique ids
            print("Waiting for {} seconds after updating the date range".format(delay_seconds))
            time.sleep(delay_seconds)

def glob_excels(excel_dir):
    return list(Path(excel_dir).glob("*.xls*"))

def clean_newlines(value):
    if isinstance(value, str):
        return value.replace(u"\n", " ")
    else:
        return value


def convert_excel(excel_dir, path_out):
    excels = glob_excels(excel_dir)
    if len(excels) != 1:
        raise ValueError(
            "Expected only one excel to "
            "be in tmp folder!, there are {}".format(excels))
    for excel in excels:
        print("converting {} into {}".format(excel, path_out))
        # saving to csc doesn't escape newlines correctly
        # soo we clean them manually
        report = (pd.read_excel(excel)
                  .applymap(clean_newlines)
                  )
        report.index.name = 'row_number'
        report.to_csv(path_out)

        # clean up!
        excel.unlink()


def main(params, datadir='/data/'):
    download_dir = '/tmp/xero_custom_reports_foo/'
    outdir = Path(datadir) / 'out/tables/'
    
    
    wd = WebDriver(headless=True, download_dir=download_dir)
    account_id = params['account_id']
    action = params['action']
    if action == 'list_reports':
        print("Listing available reports")
        with wd:
            wd.login(params['username'], params['#password'])
            reports = wd.list_reports(account_id)
            print(json.dumps(reports, indent=2))
    elif action == 'download_reports':
        print("downloading reports")
        with wd:
            wd.login(params['username'], params['#password'])
            for report in params['reports']:
                print("Downloading report", report)
                from_date = robotize_date(report.get("from_date", None))
                to_date = robotize_date(report.get("to_date", None))
                delay_seconds = report.get("delay_seconds", 15)
                try:
                    wd.download_report(report_id = report['report_id'],
                                       account_id=account_id,
                                       from_date=from_date,
                                       to_date=to_date,
                                       delay_seconds=delay_seconds)
                except Exception:
                    sc_path = '/tmp/xero_custom_reports_latest_exception.png'
                    print("Saved screenshot to", sc_path)
                    wd.driver.save_screenshot(sc_path)
                    raise
                outname = str(Path(report['filename']).stem) + '.csv'
                convert_excel(download_dir, outdir / outname)
    elif action == 'direct_url':
        print("direct_url")
        with wd:
            wd.login(params['username'], params['#password'])
            try:
                wd.direct_url(account_id=account_id,
                              url = params['direct_url'])
            except Exception:
                sc_path = '/tmp/xero_custom_reports_latest_exception.png'
                print("Saved screenshot to", sc_path)
                wd.driver.save_screenshot(sc_path)
                raise
            outname = 'report.csv'
            convert_excel(download_dir, outdir / outname)
    else:
        raise ValueError("unknown action, '{}'".format(action))


def robotize_date(dt_str):
    if dt_str is None:
        return
    cal = parsedatetime.Calendar()
    t_struct, status = cal.parse(dt_str)
    if status != 1:
        raise ValueError("Couldn't convert '{}' to a datetime".format(dt_str))
    converted = datetime.datetime(*t_struct[:6]).strftime("%-d %b %Y")
    print("converted", dt_str, "to", converted)
    return converted

if __name__ == "__main__":
    with open("/data/config.json") as f:
        cfg = json.load(f)
    try:
        main(cfg["parameters"])
    except Exception as err:
        print(err)
        traceback.print_exc()
        sys.exit(1)
