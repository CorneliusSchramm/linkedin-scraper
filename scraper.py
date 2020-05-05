 # -*- coding: utf-8 -*-

# loading required packages
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import time
from bs4 import BeautifulSoup
import csv
import re
import pandas as pd 
import numpy as np 

#session passwort und email
user_mail = input("Please input email: ")
user_password = input("Please input password: ")


# Reading the input file with all the links; all valid linked in links into url list 
masterlist = pd.read_excel("inputlist.xlsx", sheet_name='links')
urls = [e for e in masterlist["url"] if "linkedin" in e ]

# For progress tracking
num_links = len(urls)
i = 1

#Scraping and writing the output file (in utf-8 encoding)
with open("linkedin.csv", "w", newline="", encoding="utf-8") as x:
    
    #Column Headers ---
    csv.writer(x).writerow(["url","name","org","last_position","location","uni","slogan","org_url"])
    
    # using chromedriver for websraping (see exe in file directory) 
    driver = webdriver.Chrome()

    #Logging into LikedIn
    driver.get('https://www.linkedin.com/login')
    username = driver.find_element_by_name("session_key")

    username.send_keys(user_mail)
    time.sleep(0.5)
    password = driver.find_element_by_name('session_password')
    password.send_keys(user_password)
    time.sleep(0.5)
    sign_in_button = driver.find_element_by_class_name('btn__primary--large')
    sign_in_button.click()

    
    
    # Looping over all the links in the link list
    for url in urls:
        percent = round( ( (i/num_links)*100 ) ,1)
        print("Progress: "+ str( percent ) + "%" )

        # Going to a specific contact page ---
        try: 
            driver.get(url)
            #zooming out do page loads fully
            driver.execute_script("document.body.style.zoom='20%'")
            # waiting for page/javascript to load
            time.sleep(2)
            # saving html code
            page = driver.page_source
            # feeding html code to beautifulsoup
            soup = BeautifulSoup(page, 'html.parser')

            #Setting names in case of error
            name = "NA"
            org = "NA"
            slogan = "NA"
            location = "NA"
            uni = "NA"
            position = "NA"

            try:
                name = soup.find('li',class_="inline t-24 t-black t-normal break-words").text.strip()
            except:
                pass

            try:
                org = soup.find("span",class_="text-align-left ml2 t-14 t-black t-bold full-width lt-line-clamp lt-line-clamp--multi-line ember-view").text.strip()
            except:
                
                pass
            
            try: 
                slogan = soup.find("h2",class_="mt1 t-18 t-black t-normal break-words").text.strip().replace('\n', "")
                #print(slogan)           
            except:
                pass

            try:
                location = soup.find("ul", class_="pv-top-card--list pv-top-card--list-bullet mt1").text.strip().split("\n")[0]
            except:
                pass

            try: 
                uni = soup.findAll("span", class_="text-align-left ml2 t-14 t-black t-bold full-width lt-line-clamp lt-line-clamp--multi-line ember-view")[1].text.strip()
            except: 
                pass

            try:        
                if "Name des Unternehmens" not in soup.find_all('h3',class_="t-16 t-black t-bold")[0].text.strip():
                    
                    position = soup.find_all('h3',class_="t-16 t-black t-bold")[0].text.strip()

                    try:                            
                        org_url = soup.find("div", class_="display-flex flex-column full-width").find_all("a")[0]["href"]
                        org_url = "linkedin.com"+str(org_url)
                        org_url = org_url[:-1]

                        
                    except:
                        pass

                else:
                    
                    try:
                        position = soup.find_all("div",class_="pv-entity__role-container")[0].find("h3", class_="t-14 t-black t-bold").find_all("span")[1].text.strip()
                        
                        org_url = soup.find("div", class_="display-flex justify-space-between full-width").find_all("a")[0]["href"]
                        
                        org_url = "linkedin.com"+str(org_url)
                        org_url = org_url[:-1]
                    except:
                        pass
            except:
                
                pass


            #Writing CSV
            csv.writer(x).writerow([url, name,org, position, location, uni, slogan, org_url])
            i += 1

        except Exception as e:
            print(e)
            pass

    # quitting chromedriver
    driver.quit()
print("*********** DONE SCRAPING ******************\n\n")
print("*********** NOW MERGING WITH COMPANY DF AND WRITING EXCEL *****************")
# Reading the 2 dfs
comp_df = pd.read_csv("companies_sorted.csv")
li_df = pd.read_csv("linkedin.csv")

#getting relevant info from comapies data set
comp_df = comp_df[["industry","linkedin url"]]
comp_df.columns = ["industry", "org_url"]

#merging
DF = pd.merge(li_df,comp_df,on='org_url',how='left')
DF = pd.merge(masterlist,DF, on = "url", how= "left")
#saving
DF.to_excel("linkedin_with_sector.xlsx")

print("************* Programme completed *******************")