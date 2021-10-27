from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import time
import os


if not os.path.exists('output'):
    os.mkdir('output')


lib = Files()
browser_lib = Selenium()
path = f"{os.path.join(os.getcwd())}/output/"
browser_lib.set_download_directory(path+'pdf/')
x_file = 'myxlsx.xlsx'


def open_the_website(url):
    browser_lib.open_available_browser(url)


def click_dive():
    browser_lib.find_element(
        '//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]').click()


def get_agen_info():
    time.sleep(3)
    names_of_ag = browser_lib.find_elements(
        '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
    amount, departament = [], []
    for ag in names_of_ag:
        agen_info = ag.text.split("\n")
        departament.append(agen_info[0])
        amount.append(agen_info[2])
    return {"Department": departament, "Amount": amount}


def write_agen_info():
    info = get_agen_info()
    new_workbook = lib.create_workbook(path+x_file)
    new_workbook.rename_worksheet('Agencies')
    new_workbook.set_cell_value(1, 1, 'Department')
    new_workbook.set_cell_value(1, 2, 'Amount')
    new_workbook.set_cell_value(1, 3, 'Num')
    cnt = 2
    for dep in info["Department"]:
        new_workbook.set_cell_value(cnt, 1, dep)
        cnt += 1

    cnt, i = 2, 0
    for ag in info["Amount"]:
        new_workbook.set_cell_value(cnt, 2, ag)
        '''Add num to select agen'''
        new_workbook.set_cell_value(cnt, 3, i)
        cnt += 1
        i += 1
    new_workbook.save()


def select_agen():
    open_the_website("https://itdashboard.gov/")
    click_dive()
    time.sleep(3)
    new_workbook = lib.open_workbook(path+x_file)
    val = new_workbook.get_cell_value(5, 3)
    browser_lib.find_elements(
        '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]//div[@class="row top-gutter-20"]//div[@class="col-sm-12"]')[val].click()
    time.sleep(5)


def table_with_info():
    while True:
        try:
            tb_heads = browser_lib.find_element(
                '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]').find_element_by_tag_name(
                "thead").find_elements_by_tag_name("tr")[1].find_elements_by_tag_name("th")
            if tb_heads:
                break
        except Exception as e:
            print(e)
            time.sleep(1)
    headers = []
    for tb in tb_heads:
        headers.append(tb.text)

    rows, links = [], []
    while True:
        obh_to_assert = browser_lib.find_element("investments-table-object_info").text
        tb_rows = browser_lib.find_element("investments-table-object").find_element_by_tag_name(
            "tbody").find_elements_by_tag_name("tr")
        tb_links = browser_lib.find_elements('//tr[@role="row"]')

        for tb_link in tb_links[2:]:
            try:
                link = tb_link.find_element_by_tag_name('a').get_attribute("href")
            except Exception as e:
                print(e)
                link = ""
            if link:
                links.append(link)

        for tb in tb_rows:
            for tb_find in tb.find_elements_by_tag_name("td"):
                rows.append(tb_find.text)
        if browser_lib.find_element('investments-table-object_next').get_attribute(
                "class") == 'paginate_button next disabled':
            break
        else:
            browser_lib.find_element('investments-table-object_next').click()
            while True:
                if obh_to_assert != browser_lib.find_element("investments-table-object_info").text:
                    break
                time.sleep(2)

    return {'Headers': headers, 'Rows': rows, "Links": links}


def write_new_worksheet_and_down_pdf():
    new_workbook = lib.open_workbook(path+x_file)
    new_workbook.create_worksheet('All_info')
    data = table_with_info()
    cnt = 1
    for head in data["Headers"]:
        new_workbook.set_cell_value(1, cnt, head)
        cnt += 1
    cnt, i, step = 1, 2, 0
    for row in data["Rows"]:
        if step == 7:
            step, cnt = 0, 1
            i += 1
        new_workbook.set_cell_value(i, cnt, row)
        step += 1
        cnt += 1
    new_workbook.save()

    for link in data["Links"]:
        browser_lib.go_to(link)
        range_time = time.time() + 12
        while True:
            try:
                if range_time <= time.time():
                    break
                pdf_link = browser_lib.find_element('//*[contains(@id,"business-case-pdf")]//a').get_attribute("href")
                if pdf_link:
                    browser_lib.find_element('//div[@id="business-case-pdf"]').click()
                    while True:
                        try:
                            time.sleep(2)
                            if browser_lib.find_element('//div[@id="business-case-pdf"]').find_element_by_tag_name("span"):
                                time.sleep(1)
                            else:
                                break
                        except Exception as e:
                            print(e)
                            if browser_lib.find_element('//*[contains(@id,"business-case-pdf")]//a[@aria-busy="false"]'):
                                time.sleep(1)
                                break
                    break
            except Exception as e:
                print(e)
                time.sleep(1)


def main():
    try:
        open_the_website("https://itdashboard.gov/")
        click_dive()
        write_agen_info()
        select_agen()
        write_new_worksheet_and_down_pdf()
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
