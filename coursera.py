from bs4 import BeautifulSoup
import requests
from lxml import etree
from openpyxl import Workbook


def get_courses_list():
    list_of_links = []
    coursera_xml_url = "https://www.coursera.org/sitemap~www~courses.xml"
    coursera_data = requests.get(coursera_xml_url).content
    for link in etree.fromstring(coursera_data).getchildren():
        list_of_links.append(link.getchildren()[0].text.strip())
    return list_of_links


def get_course_info(link):
    course_info = []
    r = requests.get(link)
    soup = BeautifulSoup(r.content, "html.parser")
    get_header = soup.html.head.title.string
    header_div = get_header.find("|")
    header = get_header[:header_div]
    course_info.append(header)
    language_of_course = soup.find_all("div", "rc-Language")[0].next.next
    course_info.append(language_of_course)
    date = soup.find_all("div", "startdate rc-StartDateString caption-text", 'span')[0].next.next
    course_info.append(date)
    amount_of_weeks = len(soup.find_all('div', 'week'))
    course_info.append(amount_of_weeks)
    if soup.find_all('div', 'ratings-info'):
        raiting = soup.find_all('div', 'ratings-text bt3-visible-xs')[0].next.next
    else:
        raiting = "raiting is absent"
    course_info.append(raiting)
    return course_info


def output_courses_info_to_xlsx(list_of_links):
    column = 1
    row = 2
    exel_file = Workbook()
    active_exel_sheet = exel_file.active
    active_exel_sheet['A1'] = "Course name"
    active_exel_sheet['B1'] = "Language"
    active_exel_sheet['C1'] = "Start date"
    active_exel_sheet['D1'] = "Number of weeks"
    active_exel_sheet['E1'] = "Raiting"
    for link in list_of_links:
        course_info = get_course_info(link)
        for feature in course_info:
            active_exel_sheet.cell(row=row, column=column).value = feature
            exel_file.save("sample.xls")
            column += 1
        column = 1
        row += 1
    return "Thank you for using the program"


if __name__ == '__main__':
    list_of_links = []
    count = 0
    number_of_random_coursera_courses = 20
    while count != number_of_random_coursera_courses:
        list_of_links.append(get_courses_list()[count])
        count += 1
    print(output_courses_info_to_xlsx(list_of_links))
