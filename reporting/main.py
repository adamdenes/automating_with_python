#!/usr/bin/env python

import re
import time 
import requests
from selenium import webdriver
from reporting import reporting
from selenium.webdriver import Firefox, FirefoxProfile

# create a session object
session = requests.Session()
#print(session)
# create cookiejar
jar = requests.cookies.RequestsCookieJar()
#print(jar)
# dump the values seen in devtools -> network tab (cookie)
jar.set('biasValidUser', 'APPS=COM%2DTOOL&EMAIL=adam%2Edenes%40bt%2Ecom&BIASUSERID=612561487')
session.cookies = jar
#print(session.cookies)

profile = FirefoxProfile()
profile.set_preference('network.automatic-ntlm-auth.trusted-uris', 'bias2')

r = session.get('http://bias2/cmt/com-tool_x.asp')
#print(r)

browser = webdriver.Firefox(firefox_profile=profile)
#print(browser)
browser.get('http://bias2/cmt/com-tool_x.asp')
browser.find_element_by_xpath("//select[@name='s_field']/option[text()='Customer_Name']").click()


# SAVIGN THE CUSTOMERS & GROUPS TO 'cm_name, assignee_grp' variables to iterate over them
#cm_name, assignee_grp = reporting('.\\report.csv')
cm_name, assignee_grp = reporting('test.csv')

def getCustomer(customer_list):
    # EMPTY LIST THAT HOLDS THE CUSTOMER NAMES
    customer_name_list = []

    for name in customer_list:
        # SELECT THE INPUT FIELD & INSERT THE CUSTOMER's NAME
        browser.find_element_by_css_selector('input[name="s_txt"]').send_keys(name)
        # SEARCH FOR THE CUSTOMER
        browser.find_element_by_css_selector('input[type="submit"]').click()
        time.sleep(5)
        # CREATE A LIST FROM THE TABLE ROW DATA
        distributors = browser.find_elements_by_css_selector('.OrderCellTxt')
        tables = browser.find_elements_by_css_selector('.OrderForm')

        if len(distributors) > 0:
            # IF THE FIELD IS EMPTY -> THERE IS NO CUSTOMER CONTRACT
            if distributors[3].text == '':
                print('Customer_Name: -> "{}" : Contract: -> "{}" -> no contract found'.format(name, distributors[3].text))
                customer_name_list.append('NO CONTRACT FOUND')
                browser.find_element_by_css_selector('input[name="s_txt"]').clear()
            # IF THE FIELD IS NOT EMPTY 
            elif distributors[3].text:
                # CREATE A LIST FOR ALL THE TABLES IN CASE THERE IS MULTIPLE CONTRACT
                value_holder = [d.text for d in distributors]
                # CREATE A LIST WHICH SAVES EVERY 5th VALUE (THE CONTRACT)
                cst_value = [value_holder[x] for x in range(3,len(value_holder),5)] 

                cst_value_joined = ','.join(cst_value)
                regex = re.compile(r",,+", re.IGNORECASE)
                cst_value_clean = re.sub(regex, ',', cst_value_joined)

                print('Customer_Name: -> "{}" : Contract: -> "{}"'.format(name, cst_value_clean))
                customer_name_list.append(cst_value_clean)
                browser.find_element_by_css_selector('input[name="s_txt"]').clear()
        else:
            print('Customer_Name: -> customer not found')
            customer_name_list.append('NO CUSTOMER FOUND')
            browser.find_element_by_css_selector('input[name="s_txt"]').clear()

    return customer_name_list

# CALL THE FUNCTION AND INSERT THE CSV DATA (ONLY NAMES)
result = getCustomer(cm_name)

# CREATEA THE OUTPUT FILE AS CSV
output_file = 'testreport.csv'

with open(output_file, 'w') as new_csv:
    for r in range(len(result)):
        #print('{},{},{}\n'.format(cm_name[r], assignee_grp[r], result[r]))
        new_csv.write('{},{},{}\n'.format(cm_name[r], assignee_grp[r], result[r]))
    
    new_csv.close()
    print('\n%s has been successfuly created!' % (output_file))