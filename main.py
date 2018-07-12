import json
import traceback
import sys
import pandas as pd
from pathlib import Path
import time
import requests
import os
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

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

        self._account_id = None

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
        # assert we are logged in
        # TODO

    def list_reports(self):
        self.driver.get(
            "https://reporting.xero.com/{}/v2/ReportList/CustomReports?".format(
            self.account_id))
        # it's a json returned in html
        js = json.loads(self.driver.find_element_by_tag_name('pre').text)
        return {r["id"]: r["name"] for r in js['reports']}
        # print report ids and names here?

    @property
    def account_id(self):
        self.driver.get("https://reporting.xero.com/custom-reports")
        if self._account_id is not None:
            return self._account_id
        else:
            try:
                # visiting this redirects to
                # "https://reporting.xero.com/!abcde/custom-reports" where  !abcde is the account id
                self._account_id = re.search(r"(?!=.com/)!.*(?=/)", self.driver.current_url).group()
                return self._account_id
            except AttributeError:
                raise ValueError("Couldn't find account_id in {}".format(self.driver.current_url))

    def download_report(self, report_id):
        # the first string after xero.com/ is the account id
        report_download_template = (
            "https://reporting.xero.com/"
            "{account_id}/v1/Run/"
            "{report_id}?isCustom=True").format(account_id=self.account_id, report_id=report_id)

        print("getting report from ", report_download_template)
        self.driver.get(report_download_template)
        self.enable_download_in_headless_chrome()
        export_btn = self.driver.find_element_by_class_name("export-button")
        export_btn.click()
        excel_btn = [l for l in self.driver.find_elements_by_class_name("x-menu-item-link") if l.text == 'Excel'][0]
        excel_btn.click()
        # the sleeps are experimental and it might happen that the file won't be downloaded
        time.sleep(5)

        print("Report downloaded to ", glob_excels(self.download_dir))


def glob_excels(excel_dir):
    return list(Path(excel_dir).glob("*.xls*"))

def convert_excel(excel_dir, path_out):
    excels = glob_excels(excel_dir)
    if len(excels) != 1:
        raise ValueError(
            "Expected only one excel to "
            "be in tmp folder!, there are {}".format(excels))
    for excel in excels:
        print("converting {} into {}".format(excel, path_out))
        pd.read_excel(excel).to_csv(path_out, index=False, sep=u"\u0001")
        # clean up!
        excel.unlink()


def main(params, datadir='/data/'):
    download_dir = '/tmp/xero_custom_reports_foo/'
    outdir = Path(datadir) / 'out/tables/'

    wd = WebDriver(headless=True, download_dir=download_dir)
    action = params['action']
    if action == 'list_reports':
        print("Listing available reports")
        with wd:
            wd.login(params['username'], params['#password'])
            reports = wd.list_reports()
            print(json.dumps(reports, indent=2))
    elif action == 'download_reports':
        print("downloading reports")
        with wd:
            wd.login(params['username'], params['#password'])
            for report in params['reports']:
                wd.download_report(report['report_id'])
                outname = str(Path(report['filename']).stem) + '.csv'
                convert_excel(download_dir, outdir / outname)
    else:
        raise ValueError("unknown action, '{}'".format(action))



if __name__ == "__main__":
    with open("/data/config.json") as f:
        cfg = json.load(f)
    try:
        main(cfg["parameters"])
    except Exception as err:
        print(err)
        traceback.print_exc()
        sys.exit(1)
