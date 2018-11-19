from selenium import webdriver

from selenium.webdriver.common.keys import Keys

import time

import random

import urllib

from PIL import Image as IM

import xlwings as xw

import glob

import easygui

import os

import sys

reload(sys)

sys.setdefaultencoding("utf_8")

sys.getdefaultencoding()


####################################################################


browser = webdriver.Chrome(executable_path='chromedriver.exe')


def scroll_down_page():
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    small_wait()

###############################################

def excel_read():

    wb = xw.Book('Companies and Keywords.xlsx')

    Search = wb.sheets['Search']

    Result = wb.sheets['Result']

    Image = wb.sheets['Image']

    Result.range('A2:G10000').value = None

    Image.range('A2:D1000').value = None

    company_name =[]

    keywords = []

    name_count = 2

    keyword_count = 2

    while(Search.range('A' + str(name_count)).value != None):

        company_name.append(Search.range('A' + str(name_count)).value)

        name_count = name_count + 1

    while (Search.range('B' + str(keyword_count)).value != None):
        keywords.append(Search.range('B' + str(keyword_count)).value)

        keyword_count = keyword_count + 1

    return company_name,keywords


###############################################

def prioritize_keywords(Keywords):

    First_Priority = []

    Second_Priority = []

    Third_Priority = []

    prioritize = []

    for keyword in Keywords:

        if (("dev").upper() in str(keyword).upper() or "Merger".upper() in str(keyword).upper() or "M&A".upper() in str(keyword).upper()):

            First_Priority.append(str(keyword))

        if (("CSO").upper() in str(keyword).upper() or "Strategy".upper() in str(
                    keyword).upper() or "Corporate Strategy".upper() in str(keyword).upper()):

            Second_Priority.append(str(keyword))

        if (("CFO").upper() in str(keyword).upper() or "Finance".upper() in str(
                    keyword).upper() or "Financial".upper() in str(keyword).upper()):

            Third_Priority.append(str(keyword))

    check = False

    First = ','.join(str(e) for e in First_Priority)

    Second =','.join(str(e) for e in Second_Priority)

    Third = ','.join(str(e) for e in Third_Priority)

    prioritize.append(First)

    prioritize.append(Second)

    prioritize.append(Third)

    return prioritize


###############################################

def Long_wait():

    rand = random.randint(5, 8)
    browser.implicitly_wait(50)
    time.sleep(rand)


###############################################

def Medium_wait():

    rand = random.randint(3, 5)
    browser.implicitly_wait(50)
    time.sleep(rand)


###############################################

def small_wait():

    rand = random.randint(2, 3)
    browser.implicitly_wait(50)
    time.sleep(rand)


###############################################

def login(user,passw):

    browser.get('https://www.linkedin.com/')

    Medium_wait()

    username = browser.find_element_by_id("login-email")

    password = browser.find_element_by_id("login-password")

    Medium_wait()

    username.send_keys(str(user))

    small_wait()

    password.send_keys(str(passw))

    small_wait()

    browser.find_element_by_id("login-submit").click()

    Long_wait()

####################################


def get_data():


    Target_name =[]

    Target_designation = []

    Target_link =[]

    small_wait()

    titles_connect = browser.find_elements_by_xpath("//span[@class='name actor-name']//parent::span[@class='name-and-distance']//parent::span[@class='name-and-icon']//parent::h3//parent::a//parent::div//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//span[@dir='ltr']")

    names_connect = browser.find_elements_by_xpath("//span[@class='name actor-name']")

    link_connect = browser.find_elements_by_xpath("//span[@class='name actor-name']//parent::span[@class='name-and-distance']//parent::span[@class='name-and-icon']//parent::h3//parent::a")

    # # With connect Option #
    #
    # titles_connect = browser.find_elements_by_xpath("//button[@class='search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//span[@dir='ltr']")
    #
    # names_connect = browser.find_elements_by_xpath("//button[@class='search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//preceding-sibling::a[@class='search-result__result-link ember-view']//span[@class='name actor-name']")
    #
    # link_connect = browser.find_elements_by_xpath("//button[@class='search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//preceding-sibling::a[@class='search-result__result-link ember-view']")


    # # with message_lock option #
    #
    # titles_message = browser.find_elements_by_xpath("//a[@title='Send InMail']//parent::div[@class='premium-upsell-link ember-view']//parent::div[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5 link-without-visited-state']//parent::artdeco-hoverable-trigger[@class='ember-view']//parent::div[@class='ember-view']//parent::div//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//span[@dir='ltr']")
    #
    # names_message = browser.find_elements_by_xpath("//a[@title='Send InMail']//parent::div[@class='premium-upsell-link ember-view']//parent::div[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5 link-without-visited-state']//parent::artdeco-hoverable-trigger[@class='ember-view']//parent::div[@class='ember-view']//parent::div//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//preceding-sibling::a[@class='search-result__result-link ember-view']//span[@class='name actor-name']")
    #
    # link_message = browser.find_elements_by_xpath("//a[@title='Send InMail']//parent::div[@class='premium-upsell-link ember-view']//parent::div[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5 link-without-visited-state']//parent::artdeco-hoverable-trigger[@class='ember-view']//parent::div[@class='ember-view']//parent::div//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//preceding-sibling::a[@class='search-result__result-link']")
    #
    # # with message_unlock option #
    #
    # titles_umessage = browser.find_elements_by_xpath("//button[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']//span[@dir='ltr']")
    #
    # names_umessage = browser.find_elements_by_xpath("//button[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//span[@class='name actor-name']")
    #
    # link_umessage = browser.find_elements_by_xpath("//button[@class='message-anywhere-button search-result__actions--primary button-secondary-medium m5']//parent::div[@class='ember-view']//parent::div[@class='ember-view']//parent::div[@class='search-result__actions']//preceding-sibling::div[@class='search-result__info pt3 pb4 ph0']//a[@class='search-result__result-link ember-view']")


    for name, title1, link in zip(names_connect, titles_connect,link_connect):
        Target_name.append(name.text)
        Target_designation.append(title1.text)
        Target_link.append(link.get_attribute("href"))

    # for name, title, link in zip(names_message, titles_message,link_message):
    #     Target_name.append(name.text)
    #     Target_designation.append(title.text)
    #     Target_link.append(link.get_attribute("href"))
    #
    # for name, title, link in zip(names_umessage, titles_umessage,link_umessage):
    #     Target_name.append(name.text)
    #     Target_designation.append(title.text)
    #     Target_link.append(link.get_attribute("href"))

    return Target_name,Target_designation,Target_link

###############################################


def Search(company, keywords):

    search = browser.find_element_by_xpath('//input[@role="combobox"]')

    small_wait()

    company_name = 'company: "' + company + '"'

    key_words = keywords

    if ',' in keywords:

        key_words = keywords.replace(',', '" OR "')


    keyword = 'title:' + '("' + key_words + '")'

    query = company_name + " " + keyword

    search.send_keys(Keys.CONTROL,'a')

    search.send_keys(Keys.BACKSPACE)

    search.send_keys(query)

    small_wait()

    search.send_keys(u'\ue007')

    Medium_wait()

    for i in range(0, 5):

        search.send_keys(Keys.PAGE_DOWN)

        small_wait()

    for i in range(0, 5):

        search.send_keys(Keys.PAGE_UP)

        small_wait()

    Medium_wait()


###############################################


def finalizing_data(names, designations, links,count):


    wb = xw.Book('Companies and Keywords.xlsx')

    Result = wb.sheets['Result']


    for designation, name, link in zip(designations, names, links):

        if (("dev").upper() in str(designation).upper() or "Merger".upper() in str(designation).upper() or "M&A".upper() in str(designation).upper() or ("CSO").upper() in str(designation).upper() or "Strategy".upper() in str(designation).upper() or "Corporate Strategy".upper() in str(designation).upper() or ("CFO").upper() in str(designation).upper() or "Finance".upper() in str(designation).upper() or "Financial".upper() in str(designation).upper()):

            Result.range('A' + str(count)).value = str(name)

            Result.range('B' + str(count)).value = designation

            Result.range('C' + str(count)).value = str(link)

            count = count + 1

    return count

    #     designation = str(designation)
    #
    #     name = str(name)
    #
    #     Check_status = str(status)
    #
    #     if (Check_status.find(str(Prioritize[0])) or Check_status.find(str(Prioritize[1])) or Check_status.find(
    #             str(Prioritize[2]))):
    #         final_names.append(name)
    #
    #         final_designations.append(designation)
    #
    #         final_status.append(str(status))
    #
    # for name, fdesignation, status in zip(final_names, final_designations, final_status):


###############################################


def get_images(company):

    logos =[]

    Company_details = []

    Company_Name = []

    Company_links = []

    browser.get('https://www.linkedin.com/search/results/companies/?keywords=company%3A%20 "' + company +
                '" &origin=SWITCH_SEARCH_VERTICAL')

    Long_wait()

    search = browser.find_element_by_xpath('//input[@role="combobox"]')

    for i in range(0, 5):
        search.send_keys(Keys.PAGE_DOWN)
        small_wait()

    for i in range(0, 5):
        search.send_keys(Keys.PAGE_UP)
        small_wait()


    company_name = browser.find_elements_by_xpath("//h3[@class='search-result__title t-16 t-black t-bold']")

    small_wait()

    company_logo = browser.find_elements_by_xpath("//div[@class='ivm-view-attr__img--centered  EntityPhoto-square-4 ember-view']")

    small_wait()

    company_detail = browser.find_elements_by_xpath("//div[@class='ivm-view-attr__img--centered  EntityPhoto-square-4 ember-view']//parent::div[@class='display-flex ember-view']//parent::div[@class='ivm-image-view-model ember-view']//parent::figure[@class='search-result__image']//parent::a[@class='search-result__result-link ember-view']//parent::div[@class='search-result__image-wrapper']//following-sibling::div[@class='search-result__info pt3 pb4 pr0']//p[@class='subline-level-1 t-14 t-black t-normal search-result__truncate']")

    small_wait()

    company_links = browser.find_elements_by_xpath("//h3[@class='search-result__title t-16 t-black t-bold']//parent::a[@class='search-result__result-link ember-view']")

    for logo,detail,name,links in zip(company_logo,company_detail,company_name,company_links):

        if (str(company).upper() in str(name.text).upper()):

            logos.append(logo.get_attribute("style"))

            Company_details.append(detail.text)

            Company_Name.append(name.text)

            Company_links.append(links.get_attribute("href"))

    return logos,Company_details,Company_Name,Company_links

###############################################



def extract_image_link(links):

    image_links = []

    for link in links:

        link = str(link)

        link = (link[link.find('"')+1:link.rfind('"')])

        image_links.append(link)

    return image_links
    #
    # count = 1
    #
    # for company in Companies:
    #
    #     for priority in Prioritize:
    #
    #         Search(str(company), str(priority))
    #
    #         names, designations, links = get_data()
    #
    #         count = finalizing_data(names, designations, links,count)
    #
    #
    # get_images(Companies)


###############################################

def finalize_image(Companies):

    image_count = 2

    wb = xw.Book('Companies and Keywords.xlsx')

    Image = wb.sheets('Image')


    for company in Companies:

        images, details, names , links = get_images(company)

        image_link = extract_image_link(images)


        for image, detail,name,link in zip(image_link, details,names,links):

            Image.range("A" + str(image_count)).value = name

            Image.range("B" + str(image_count)).value = detail

            Image.range("C" + str(image_count)).value = link

            image_count = image_count + 1


        for link,name in zip(links,names):


            if '|' in name or '/' in name or ':' in name or '*' in name or '?' in name or '"' in name or '<' in name or '>' in name:

                name = str(name).replace('|',' ').replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ')


            browser.get(link)

            Medium_wait()

            image = browser.find_element_by_xpath("//img[@class='lazy-image org-top-card-module__logo loaded']")

            small_wait()

            image = image.get_attribute("src")

            small_wait()

            urllib.urlretrieve('http:'+str(image)[str(image).find(':')+1:],"Images/"+str(name)+".jpg")

            small_wait()

            im1 = IM.open("Images/"+str(name)+".jpg")

            im_small = im1.resize((93, 37), IM.ANTIALIAS)

            im_small.convert("RGB").save("Images/"+str(name)+".jpg")


###############################################

def total_employees():

    wb = xw.Book('Companies and Keywords.xlsx')

    Image = wb.sheets('Image')

    Company_links = []

    company_count = 2

    link_count = 2

    while (Image.range('C' + str(company_count)).value != None):

        Company_links.append(Image.range('C' + str(company_count)).value)

        company_count = company_count + 1

    for link in Company_links:

        browser.get(link)

        Medium_wait()

        employees = browser.find_element_by_xpath("//strong[contains(text(),'employees')]")

        small_wait()

        employees = employees.text

        employees = str(employees)[employees.find("all")+3:employees.rfind("on")-1]

        Image.range('D' + str(link_count)).value = employees

        link_count = link_count + 1


###############################################


def remove_duplicates():

    wb = xw.Book('Companies and Keywords.xlsx')

    Result = wb.sheets('Result')

    Name = []

    Title = []

    Profile_link = []

    data_count = 2



    Final_Name = []

    Final_Title = []

    Final_Profile_link = []

    insert_count = 2


    while (Result.range('A' + str(data_count)).value != None):

        Name.append(Result.range('A' + str(data_count)).value)

        Result.range('A' + str(data_count)).value = None

        Title.append(Result.range('B' + str(data_count)).value)

        Result.range('B' + str(data_count)).value = None

        Profile_link.append(Result.range('C' + str(data_count)).value)

        Result.range('C' + str(data_count)).value = None

        data_count = data_count + 1

    for name,title,link in zip(Name,Title,Profile_link):

        if name not in Final_Name:

            Final_Name.append(name)

            Final_Title.append(title)

            Final_Profile_link.append(link)

    for name, title, link in zip(Final_Name, Final_Title, Final_Profile_link):

        Result.range('A' + str(insert_count)).value = name

        Result.range('B' + str(insert_count)).value = title

        Result.range('C' + str(insert_count)).value = link

        insert_count = insert_count + 1

###############################################

def employees_companies():

    wb = xw.Book('Companies and Keywords.xlsx')

    Result = wb.sheets('Result')

    Image = wb.sheets('Image')

    Profile_link = []

    company_name = []

    company_details = []

    company_link = []

    company_employees = []

    profile_count = 2

    company_count = 2

    match_count = 2

    while (Image.range('C' + str(company_count)).value != None):

        company_name.append(Image.range('A' + str(company_count)).value)

        company_details.append(Image.range('B' + str(company_count)).value)

        company_link.append(Image.range('C' + str(company_count)).value)

        company_employees.append(Image.range('D' + str(company_count)).value)

        company_count = company_count + 1



    while (Result.range('C' + str(profile_count)).value != None):

        Profile_link.append(Result.range('C' + str(profile_count)).value)

        profile_count = profile_count + 1



    for link in Profile_link:

        browser.get(link)

        Medium_wait()

        search = browser.find_element_by_xpath('//input[@role="combobox"]')

        for i in range(0, 5):

            search.send_keys(Keys.PAGE_DOWN)

            small_wait()

        for i in range(0, 5):

            search.send_keys(Keys.PAGE_UP)

            small_wait()

        companies = browser.find_elements_by_xpath("//a[@data-control-name='background_details_company']")


        for name,detail,link,emp in zip(company_name,company_details,company_link,company_employees):

            for company in companies:

                if  company.get_attribute("href") in link:

                    Result.range('D' + str(match_count)).value = name

                    Result.range('E' + str(match_count)).value = detail

                    Result.range('F' + str(match_count)).value = link

                    Result.range('G' + str(match_count)).value = emp

        match_count = match_count + 1


###############################################

def company_logos():

    wb = xw.Book('Companies and Keywords.xlsx')

    Image = wb.sheets['Image']

    name_count = 1

    images = glob.glob("Images/*.jpg")

    for company in Image.range("A:A").value:

        if company != None and company != "Company_name":

            rng = Image.range('E'+str(name_count))

            name_count = name_count + 1

            for image in images:

                image = image[int(image.find('\\')) + 1:int(image.rfind('.'))]

                if company == image:

                    Image.pictures.add(os.getcwd()+"\\" +"Images"+"\\"+str(image)+".jpg",top =rng.top + 5 ,left =rng.left + 5 )

###############################################

if __name__ == "__main__":

    names = []

    designations = []

    links = []

    count = 2

    Companies, Keywords = excel_read()

    Prioritize = prioritize_keywords(Keywords)

    wb = xw.Book('Companies and Keywords.xlsx')

    Credentials = wb.sheets['Credentials']

    login(str(Credentials.range('A2').value),str(Credentials.range('B2').value))


    for company in Companies:

        first_priority_count = 1

        iteration_count = 1

        next_found = True

        for keyword in Prioritize:

            Search(company,keyword)

            if first_priority_count == 1:

                while iteration_count < 4 and next_found == True:

                    search = browser.find_element_by_xpath('//input[@role="combobox"]')

                    for i in range(0, 5):
                        search.send_keys(Keys.PAGE_DOWN)

                        small_wait()

                    for i in range(0, 5):
                        search.send_keys(Keys.PAGE_UP)

                        small_wait()

                    names, designations, links = get_data()

                    count = finalizing_data(names, designations, links,count)

                    try:

                        browser.find_element_by_xpath("//button[@class='next']")

                        next_found = True

                    except:

                        next_found = False

                    if next_found == True:

                        next = browser.find_element_by_xpath("//button[@class='next']")

                        browser.execute_script("arguments[0].click();", next)

                        iteration_count = iteration_count + 1

                        next_found = True

                    else:

                        next_found = False

                first_priority_count = first_priority_count + 1

            else:

                search = browser.find_element_by_xpath('//input[@role="combobox"]')

                for i in range(0, 5):
                    search.send_keys(Keys.PAGE_DOWN)

                    small_wait()

                for i in range(0, 5):
                    search.send_keys(Keys.PAGE_UP)

                    small_wait()

                names, designations, links = get_data()

                count = finalizing_data(names, designations, links, count)

    remove_duplicates()

    finalize_image(Companies)

    total_employees()

    employees_companies()

    company_logos()

    easygui.msgbox("Data has been scraped successfully")




