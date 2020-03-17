# -*- coding: utf-8 -*-
"""

@author: Muhammed Ali KOCABEY
"""

from selenium import webdriver
from win32api import GetSystemMetrics
from selenium.common.exceptions import ElementNotVisibleException
import pandas as pd





driver_path = "C:\selenium_chrome_driver\chromedriver"
browser = webdriver.Chrome(executable_path=driver_path)

openSite = browser.get("https://www.microsoft.com/en-us/store/most-popular/apps/pc")

screenWidth = GetSystemMetrics(0)
screenHeight = GetSystemMetrics(1)

browser.set_window_size(screenWidth-5, screenHeight-5)

closeSelection = browser.find_element_by_xpath("//button[@id='R1MarketRedirect-close']")
closeSelection.click()


# App Area Class   =     m-channel-placement-item
# App Name Class   =     c-subheading-6
# App Star Class   =     c-rating
# App Price Class  =     c-price
# App Image Class  =     c-channel-placement-image
# App Review Class =     x-screen-reader 


app_List = list()






pageLinkActive = True


app = {
       "app_Name" : "",
       "app_Star" : "",
       "app_Price" : "",
       "app_Review" : ""
       }




while pageLinkActive:
    
    try:
        main_content = browser.find_elements_by_xpath("//div//div[@class='c-group f-wrap-items context-list-page']//div[@class='m-channel-placement-item']")
        
        for items in main_content:
            app_name = items.find_element_by_class_name("c-subheading-6")
            app_name_content = app_name.get_attribute("innerHTML")
            app["app_Name"] = app_name_content
            
            
            
            try:
                star_score = items.find_element_by_class_name("c-rating")
                star_score_content = star_score.get_attribute("data-value")
                app["app_Star"] = star_score_content
            
            except:
                app["app_Star"] = "NULL"
            
            
            
            
            
            price = items.find_element_by_xpath("//div[@class='c-price']/span[1]")
            price_content = price.get_attribute("innerHTML")
            price_content = price_content.strip()
            app["app_Price"] = price_content
                              
            
                              
            review_datas = items.find_elements_by_class_name("x-screen-reader")
        
            try:
                review_content = (review_datas[2].get_attribute("innerHTML"))
                review_content_split = review_content.split(" ")
                review_content_clean = review_content_split[2]
                app["app_Review"] = review_content_clean
                
            except:
                app["app_Review"] = "NULL"
            
            
            app_List.append(app.copy())
        
        
        
        # Last page control
        pageLink = browser.find_element_by_xpath("//ul[@class='m-pagination']//li//a[@class='c-glyph' and @aria-label='next page']")
        
        
        try:
            controlPageLink = browser.find_element_by_xpath("//ul[@class='m-pagination']//li[@class='f-hide']//a[@aria-label='next page']")
            pageLinkActive = False
            
        except:
            pageLink.click()
        
        
    except:
        closeDialog = browser.find_element_by_xpath("//div[@class='sfw-dialog']//div[@class='c-glyph glyph-cancel']")
        closeDialog.click()
        
        
# modDfObj = dfObj.append(listOfSeries , ignore_index=True)
        
app_DataFrame = pd.DataFrame.from_dict(app_List)
app_DataFrame.to_csv('Microsoft_Common_Apps.csv', sep=',', encoding='utf-8')
