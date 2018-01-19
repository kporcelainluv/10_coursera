from bs4 import BeautifulSoup
import requests
from lxml import etree
from openpyxl import Workbook
import sys
import random


def get_courses_links(url):
    list_of_links = []
    coursera_data = request_page_info(url)
    for link in etree.fromstring(coursera_data).getchildren():
        list_of_links.append(link.getchildren()[0].text.strip())
    return list_of_links


def request_page_info(link):
    webpage = requests.get(link).content
    return webpage


def get_course_info(link):
    course_info = {}
    soup = BeautifulSoup(request_page_info(link), "html.parser")
    course_name = soup.find("h1", attrs={"class": "title display-3-text"}).text
    course_info["name"] = course_name
    course_lang = soup.find("div", attrs={"class": "rc-Language"}).text
    course_info["language"] = course_lang
    starting_date = (soup.find("div",
                               attrs={"class": "startdate rc-StartDateString caption-text"}).text).split()
    course_info["date"] = " ".join(starting_date[1:])
    course_info["weeks"] = len(soup.find_all("div", "week"))
    rating_tr = soup.find("div", attrs={"class": "ratings-text bt3-hidden-xs"})
    if rating_tr is None:
        course_info["rating"] = None
    else:
        average_user_rating = rating_tr.find("span").text.split()
        course_info["rating"] = average_user_rating[-1]
    return course_info


def create_wb():
    wb = Workbook()
    active_exel_sheet = wb.active
    head_line = ["Course name",
                 "Language",
                 "Start Date",
                 "Duration, weeks",
                 "Rating"]
    active_exel_sheet.append(head_line)
    return wb, active_exel_sheet


def append_info_to_wb(list_of_links, wb, active_exel_sheet):
    for link in list_of_links:
        dict_of_course_info = get_course_info(link)
        active_exel_sheet.append([
            dict_of_course_info["name"],
            dict_of_course_info["language"],
            dict_of_course_info["date"],
            dict_of_course_info["weeks"],
            dict_of_course_info["rating"]
        ])
    return wb


if __name__ == "__main__":
    if len(sys.argv) > 1:
        url = "https://www.coursera.org/sitemap~www~courses.xml"
        xls_filename = sys.argv[1]
        num_of_courses = 20
        list_of_links = random.sample(get_courses_links(url), num_of_courses)
        wb, active_exel_sheet = create_wb()
        wb = append_info_to_wb(list_of_links, wb, active_exel_sheet)
        wb.save(xls_filename)
    print("Done! Thank you for using the program.")
