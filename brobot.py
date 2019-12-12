from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd

def fillInOaInfo(dataFrame, oaList, username, password, url = "http://10.0.3.206:8168/GlobalOA/"):
    def waitForLoadingDialog(second = "0.5"):
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
        try:
            search_elem = driver.find_element_by_css_selector(f"#resultsPane td>a[href*='{oaId}']")
        except NoSuchElementException:
            return f"don't find {oaId}"
        search_elem.click()
    
        driver.switch_to.frame(driver.find_element_by_id(oaId))
        
        subject = driver.find_element_by_css_selector('span#txtSubject.fieldText').text
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "OA_DESC"] = subject
        
        affectedSite = driver.find_element_by_css_selector('div.fieldBodyCoulmnLeft>span.privilegeStatus0>span.privilegeStatus5>div.fieldSubBlock>span.privilegeIndividual>div>span.fieldText').text
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "SITE"] = affectedSite

        dueDate = driver.find_element_by_css_selector('input#setRequestDateX').get_attribute('value').replace("-","/")
        dataFrame.loc[dataFrame["OA_NO"] == oaId, "DUE_DATE"] = dueDate
        
        driver.switch_to.default_content()
        
        #close oa tab to prevent that too many tabs to open new tab, but there might not be a tab number limit now?
        #driver.find_element_by_css_selector("span#dijit_layout__TabButton_6 + span.closeImage").click()
        
        return f"{oaId} update successfully"
        
    
    driver = webdriver.Chrome(executable_path="chromedriver.exe") # Use Chrome
    
    preparePage(driver, url)
        
    processInfo = ""
    for oaId in oaList:
        processInfo += getOaInfo(driver, dataFrame, oaId) + "\n"
    driver.close()
    return processInfo
    
