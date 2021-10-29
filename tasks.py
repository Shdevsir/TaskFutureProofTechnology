from posixpath import expanduser
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os


if not os.path.exists('output'):
    os.mkdir('output')


pdf = PDF()
lib = Files()
browser_lib = Selenium()
path = f"{os.path.join(os.getcwd())}/output/"
browser_lib.set_download_directory(path)
x_file = 'myxlsx.xlsx'
column_length = {}
limit = 6


def open_the_website(url):
    browser_lib.open_available_browser(url)


def click_dive():
    browser_lib.find_element(
        '//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]').click()


def get_agen_info():
    browser_lib.wait_until_element_is_enabled(
        '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
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
    new_workbook = lib.open_workbook(path+x_file)
    val = new_workbook.get_cell_value(5, 3)
    browser_lib.wait_until_element_is_enabled(
        '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
    browser_lib.find_elements(
        '//div[@class="row top-gutter-20"]//div[@class="col-sm-12"]')[val].click()


def table_with_info():
    browser_lib.wait_until_element_is_visible(
        '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]', timeout=20)
    tb_heads = browser_lib.find_element(
        '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]').find_element_by_tag_name(
        "thead").find_elements_by_tag_name("tr")[1].find_elements_by_tag_name("th")
    headers = []
    for tb in tb_heads:
        headers.append(tb.text)

    rows, links = [], []
    while True:
        browser_lib.wait_until_element_is_enabled('//table[@class="datasource-table usa-table-borderless dataTable no-footer"]')
        obh_to_assert = browser_lib.find_element("investments-table-object_info").text
        tb_rows = browser_lib.find_element("investments-table-object").find_element_by_tag_name(
            "tbody").find_elements_by_tag_name("tr")
        tb_links = browser_lib.find_elements('//tr[@role="row"]')

        for tb_link in tb_links[2:]:
            try:
                link = tb_link.find_element_by_tag_name('a').get_attribute("href")
            except Exception as e:
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

    return {'Headers': headers, 'Rows': rows, "Links": links}


def write_new_worksheet_and_download_pdf(limit_pdf_file):
    '''
    The limit is related to specifics of the deployment to
    https://cloud.robocorp.com/
    (file upload limit 5Mb)
    '''
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
    column_length['len'] = i - 1
    new_workbook.save()
    for link in data["Links"]:
        limit_pdf_file -= 1
        browser_lib.go_to(link)
        pdf_link = browser_lib.wait_until_element_is_visible('//*[contains(@id,"business-case-pdf")]//a')
        pdf_link = browser_lib.find_element('//*[contains(@id,"business-case-pdf")]//a').get_attribute("href")
        if pdf_link:
            browser_lib.find_element('//div[@id="business-case-pdf"]').click()
            while True:
                try:
                    if browser_lib.find_element('//div[@id="business-case-pdf"]').find_element_by_tag_name("span"):
                        continue
                    else:
                        break
                except Exception as e:
                    if browser_lib.find_element('//*[contains(@id,"business-case-pdf")]//a[@aria-busy="false"]'):
                        break
            if limit_pdf_file == 0:
                break


def compare_data(limit_pdf_file):
    '''
    The limit is related to specifics of the deployment to
    https://cloud.robocorp.com/
    (file upload limit 5Mb)
    '''
    new_workbook = lib.open_workbook(path+x_file)

    for cnt in range(column_length['len']):
        limit_pdf_file -= 1
        data_uii = new_workbook.get_cell_value(cnt+2, 1)
        data_name = new_workbook.get_cell_value(cnt+2, 3)
        pdf_info = get_pdf_info(data_uii)
        if data_uii == pdf_info["UII"] and data_name == pdf_info["Name"]:
            print(f'UII and Investment Title compare, colum: {cnt+2}')
        else:
            print(f'UII or Investment Title not compare, colum: {cnt+2}')
        cnt += 1
        if limit_pdf_file == 0:
            break


def get_pdf_info(name_pdf):
    data = pdf.get_text_from_pdf(path+name_pdf+'.pdf', 1)
    list_of_data = data[1].split()
    i = 0
    index = []
    for c in list_of_data:
        if c == "Investment:":
            index.append(i+1)
        elif c == "Unique":
            index.append(i)
        elif c == "(UII):":
            index.append(i+1)
        i += 1

    list_of_data[(index[1]-1)] = list_of_data[(index[1]-1)].replace("2.", "")
    name = ' '.join(list_of_data[index[0]:index[1]])
    list_of_data[index[2]] = list_of_data[index[2]].replace("Section", "")
    uii = list_of_data[index[2]]
    return {'Name': name, 'UII': uii}


def main():
    try:
        open_the_website("https://itdashboard.gov/")
        click_dive()
        write_agen_info()
        select_agen()
        write_new_worksheet_and_download_pdf(limit)
        compare_data(limit)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
