from bs4 import BeautifulSoup
import requests
from lxml import etree
from openpyxl import Workbook
import sys
import random


def get_courses_links(coursera_data):
    list_of_links = []
    for link in etree.fromstring(coursera_data).getchildren():
        list_of_links.append(link.getchildren()[0].text.strip())
    return list_of_links


def fetch_page_info(link):
    webpage = requests.get(link).content
    return webpage


def get_course_info(course_html):
    course_info = {}
    soup = BeautifulSoup(course_html, "html.parser")

    course_name = soup.find("h1", attrs={"class": "title display-3-text"}).text
    course_info["name"] = course_name

    course_lang = soup.find("div", attrs={"class": "rc-Language"}).text
    course_info["language"] = course_lang

    starting_date_attrib = "startdate rc-StartDateString caption-text"
    starting_date = soup.find("div", attrs={"class": starting_date_attrib})
    starting_date = starting_date.text.split()
    course_info["date"] = " ".join(starting_date[1:])

    course_info["weeks"] = len(soup.find_all("div", "week"))

    rating_attrib = soup.find("div", attrs={"class": "ratings-text bt3-hidden-xs"})

    if rating_attrib is None:
        course_info["rating"] = None
    else:
        rating = rating_attrib.find("span").text.split()
        course_info["rating"] = rating[-1]
    return course_info


def output_info_to_excel(courses_info):
    courses_workbook = Workbook()
    active_exel_sheet = courses_workbook.active
    head_line = ["Course name",
                 "Language",
                 "Start Date",
                 "Duration, weeks",
                 "Rating"]
    active_exel_sheet.append(head_line)

    for course in courses_info:
        active_exel_sheet.append([
            course["name"],
            course["language"],
            course["date"],
            course["weeks"],
            course["rating"]
            ])
    return courses_workbook


if __name__ == "__main__":
    if len(sys.argv) == 1:
        exit("Enter a name for xls file")
    xls_filename = sys.argv[1]
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    num_of_courses = 20
    coursera_courses_links = fetch_page_info(url)
    list_of_links = random.sample(
        get_courses_links(coursera_courses_links),
        num_of_courses
    )
    courses_info = []

    for course_url in list_of_links:
        course_html = fetch_page_info(course_url)
        dict_of_course_info = get_course_info(course_html)
        courses_info.append(dict_of_course_info)
    courses_workbook = (output_info_to_excel(courses_info))
    courses_workbook.save(xls_filename)
    print("Done!")
