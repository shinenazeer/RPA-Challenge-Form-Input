from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

browser = Selenium()
http = HTTP()
excel = Files()
URL = "http://rpachallenge.com/"


def download_excel_data():
    http.download("http://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx",overwrite=True)    


def open_rpa_challenge_website():
    browser.open_available_browser(URL)
    browser.maximize_browser_window()
    browser.click_button("Start")


def get_data_from_excel():
    excel.open_workbook("challenge.xlsx")
    tabledata = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()
    return tabledata


def fill_the_form():
    table = get_data_from_excel()
    for row in table:
        browser.input_text('css:input[ng-reflect-name="labelFirstName"]', row["First Name"])
        browser.input_text('css:input[ng-reflect-name="labelLastName"]', row["Last Name"])
        browser.input_text('css:input[ng-reflect-name="labelRole"]', row["Role in Company"])
        browser.input_text('css:input[ng-reflect-name="labelCompanyName"]', row["Company Name"])
        browser.input_text('css:input[ng-reflect-name="labelPhone"]', row["Phone Number"])
        browser.input_text('css:input[ng-reflect-name="labelAddress"]', row["Address"])
        browser.input_text('css:input[ng-reflect-name="labelEmail"]', row["Email"])
        browser.click_button("Submit")


def take_result_screenshot():
    browser.wait_until_page_contains_element("css:div.congratulations")
    browser.capture_element_screenshot("css:div.congratulations", filename="result.png")


def main():
    download_excel_data()
    open_rpa_challenge_website()
    fill_the_form()
    take_result_screenshot()
    browser.close_all_browsers()


if __name__ == "__main__":
    main()


