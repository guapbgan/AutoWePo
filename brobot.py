from selenium import webdriver
import time
import pandas as pd

def fillInOaInfo(dataFrame, oaList, username, password, url = "http://10.0.3.206:8168/GlobalOA/"):
    def waitForLoadingDialog(second = "1"):
        while(True):
            if driver.find_element_by_css_selector('div#loadingdialog').value_of_css_property("visibility") == "visible":
                time.sleep(1)
            else:
                break    
    
    def preparePage(driver, url):
        driver.get(url)
        search_elem = driver.find_element_by_css_selector("#username")
        search_elem.send_keys(username)
        search_elem = driver.find_element_by_css_selector("#password")
        search_elem.send_keys(password)
        search_elem = driver.find_element_by_css_selector("#loginBtn")
        search_elem.click()
        waitForLoadingDialog()
        
    
    def getOaInfo(driver, dataFrame, oaId):
        search_elem = driver.find_element_by_css_selector("#searchBox")
        search_elem.clear()
        search_elem.send_keys(oaId)
        search_elem.send_keys(u'\ue007') #press enter

        waitForLoadingDialog()
        search_elem = driver.find_element_by_css_selector(f"td>a[href*='{oaId}']")
        search_elem.click()
    
        driver.switch_to.frame(driver.find_element_by_id(oaId))
        
        subject = driver.find_element_by_css_selector('span#txtSubject.fieldText').text
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "OA_DESC"] = subject
        
        affectedSite = driver.find_element_by_css_selector('div.fieldBodyCoulmnLeft>span.privilegeStatus0>span.privilegeStatus5>div.fieldSubBlock>span.privilegeIndividual>div>span.fieldText').text
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "SITE"] = affectedSite

        dueDate = driver.find_element_by_css_selector('input#setRequestDateX').get_attribute('value').replace("-","/")
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "DUE_DATE"] = dueDate
        
        driver.switch_to.default_content()        
        
    
    driver = webdriver.Chrome(executable_path="chromedriver.exe") # Use Chrome
    
    preparePage(driver, url)
        
    for oaId in oaList:
        getOaInfo(driver, dataFrame, oaId)
    driver.close()
    
