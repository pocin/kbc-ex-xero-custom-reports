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
import maya
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

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
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36")
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

    def login(self, username, password):
        self.driver.get("https://login.xero.com")

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
                 "The title says {}. Check your advertiser ID!").format(title))

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
            wait_for_download=10):
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
        self.update_date_range(from_date, to_date)

        # click export btn so that excel btn is rendered
        export_btn = self._locate_export_button()
        export_btn.click()
        time.sleep(1)

        excel_btn = self._locate_export_to_excel_button()
        excel_btn.click()
        # the sleeps are experimental and it might happen that the file won't be downloaded
        time.sleep(wait_for_download)

        print("Report downloaded to ", glob_excels(self.download_dir))

    def _locate_export_button(self):
        for btn in  self.driver.find_elements_by_tag_name('button'):
            if btn.get_attribute('data-automationid') == 'report-toolbar-export-button':
                return btn
        raise KeyError("Couldn't find export menu. "
                       "The underlying html/css probably changed and the code needs to be adjusted")
    def _locate_export_to_excel_button(self):
        for btn in self.driver.find_elements_by_tag_name('button'):
            if btn.get_attribute('data-automationid') == 'report-toolbar-export-excel-menuitem--body':
                return btn

        raise KeyError("Couldn't find export button. "
                       "The underlying html/css probably changed and the code needs to be adjusted")

    def update_date_range(self, from_time: Union[str, None], until_time: Union[str, None]):
        # update From field
        if from_time:
            from_input_field = self.driver.find_element_by_id("dateFieldFrom-inputEl")
            time.sleep(1)
            from_input_field.send_keys(
                Keys.BACKSPACE * len(from_input_field.get_attribute("value")))

            from_input_field.send_keys(from_time)
            # press the update button

        if until_time:
            # update To field
            until_input_field = self.driver.find_element_by_id("dateFieldTo-inputEl")
            time.sleep(1)
            # clear the input field
            until_input_field.send_keys(
                Keys.BACKSPACE * len(until_input_field.get_attribute("value")))

            # input the datetime
            until_input_field.send_keys(until_time)

            # take the first div that satisfies the condition
            update_btn = next(filter(
                lambda div: div.get_attribute("data-automationid") == "date-toolbar-update-button",
                self.driver.find_elements_by_tag_name("div")))
            update_btn.click()
            time.sleep(7)
            # wait 5 seconds to the report is, hopefully, updated

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
    action = params['action']
    account_id = params['account_id']
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
                try:
                    wd.download_report(report['report_id'],
                                       account_id=account_id,
                                       from_date=from_date,
                                       to_date=to_date,
                                       wait_for_download=report.get('download_timeout', 10))
                except Exception:
                    sc_path = '/tmp/xero_custom_report_latest_exception.png'
                    print("Saved screenshot to", sc_path)
                    wd.driver.save_screenshot(sc_path)
                    raise
                outname = str(Path(report['filename']).stem) + '.csv'
                convert_excel(download_dir, outdir / outname)
    else:
        raise ValueError("unknown action, '{}'".format(action))


def robotize_date(dt_str):
    if dt_str is None:
        return
    converted = maya.when(dt_str).datetime().strftime("%d %b %Y")
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
