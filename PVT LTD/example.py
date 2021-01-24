#download image
import selenium
#import PIL
from PIL import Image
from selenium import webdriver

#death by captcha
import deathbycaptcha

# excel
import xlrd 
import xlwt
import openpyxl
import time
from openpyxl import workbook,Workbook
from time import sleep
from datetime import date

today = date.today()

  
loc = ("1.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
cid=[]
client_array=[]
email_array=[]
company_array=[]
roc_array=[]
doi_array=[]
business_array=[]
#print(1)

for i in range(0,sheet.nrows): 
    cid.append(sheet.cell_value(i, 0))

wb = Workbook() 

ws = wb.active
i=0
while i<len(cid):
        ##########################################################
    # Download image code
    ##########################################################

    def get_captcha(driver, element, path):
        # now that we have the preliminary stuff out of the way time to get that image :D
        location = element.location
        size = element.size
        # saves screenshot of entire page
        driver.save_screenshot(path)

        # uses PIL library to open image in memory
        image = Image.open(path)

        left = location['x']
        top = location['y'] 
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']

        image = image.crop((left, top, right, bottom))  # defines crop points
        image.save(path, 'png')  # saves new cropped image


    driver = webdriver.Chrome(executable_path=r'C:\Users\MCA Scrapper\Downloads\chromedriver_win32\chromedriver.exe')
    try:
        driver.get("http://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do")
    except:
        i=i-1
        print("Error1")
        driver.quit()
        sleep(10)
        continue
    # change frame
    # driver.switch_to.frame("displayCaptcha")

    # download image/captcha
    img = driver.find_element_by_xpath(".//*[@id='captcha']")
    get_captcha(driver, img, "RandomTxt.jpg")



    # driver = webdriver.Chrome(executable_path=r'C:\Users\Ansh\Downloads\chromedriver_win32\chromedriver.exe')
    # driver.get("http://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do")

    # change frame
    # driver.switch_to.frame("displayCaptcha")

    # download image/captcha
    # img = driver.find_element_by_xpath(".//*[@id='captcha']")
    # print(img)
    # get_captcha(driver, img, "RandomTxt.jpg")



    ########################################################
    # Death by captcha start
    ########################################################

    # Put your DBC account username and password here.
    # Use deathbycaptcha.HttpClient for HTTP API.
    client = deathbycaptcha.SocketClient('username', 'password')
    try:
        balance = client.get_balance()

        # Put your CAPTCHA file name or file-like object, and optional
        # solving timeout (in seconds) here:
        captcha = client.decode("RandomTxt.jpg")
        if captcha:
            # The CAPTCHA was solved; captcha["captcha"] item holds its
            # numeric ID, and captcha["text"] item its text.
            print ("CAPTCHA %s solved: %s" % (captcha["captcha"], captcha["text"]))

            if ...:  # check if the CAPTCHA was incorrectly solved
                client.report(captcha["captcha"])
    except deathbycaptcha.AccessDeniedException:
        # Access to DBC API denied, check your credentials and/or balance
        print ("error: Access to DBC API denied, check your credentials and/or balance")


    ########################################################
    # login
    ########################################################

    username = driver.find_element_by_id("companyID")
    username.clear()
    username.send_keys(cid[i])

    captchatext = driver.find_element_by_id("userEnteredCaptcha")
    captchatext.clear()
    captchatext.send_keys(captcha["text"].lower())
    try:
        driver.find_element_by_id("companyLLPMasterData_0").click()
    except:
        i=i-1
        print("Error2")
        sleep(10)
        driver.quit()
        continue
    

        ########################################################
        # fetch details
        ########################################################
    try:
        client_name = driver.find_element_by_xpath("//*[@id='resultTab6']/tbody/tr[2]/td[2]").text
                # for i in email_id:
                #    print(i.text)
                # print(client_name)
        #client_array.append(client_name)

        company_name = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[2]/td[2]").text
                # for i in email_id:
                #    print(i.text)
                # print(email_id)
        #company_array.append(company_name)

        ROC_code = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[3]/td[2]").text
                # for i in email_id:
                #    print(i.text)
                # print(email_id)
        #roc_array.append(ROC_code)

        doi = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[11]/td[2]").text
                    # for i in email_id:
                    #    print(i.text)
                    # print(email_id)
        #doi_array.append(doi)    

        #business = driver.find_element_by_xpath(".//*[@id='resultTab3']/tbody/tr[12]/td[2]").text
                    # for i in email_id:
                    #    print(i.text)
                    # print(email_id)
        #business_array.append(business)  

        email_id = driver.find_element_by_xpath("//*[@id='resultTab1']/tbody/tr[14]/td[2]").text
                    # for i in email_id:
                    #    print(i.text)
                    # print(email_id)
        #email_array.append(email_id)
    except:
        i=i-1
        print(client_name)
        print("Error3")
        sleep(10)
        driver.quit()
        continue

    ws.append([client_name,company_name,email_id,doi,ROC_code])
    wb.save('PVT data.xls') 
  
            ########################################################
            # close browser window
            ########################################################
    print(i)
    driver.quit()
            # Time in seconds
    if(i%2==0):
        sleep(3)
    i=i+1    

# for i in range(0,len(email_array)):
#     print(client_array[i])
#     print(company_array[i])
#     print(email_array[i])
#     print(business_array[i])
#     print(doi_array[i])
#     print(roc_array[i])


# Workbook is created 
# wb = Workbook() 

# ws = wb.active
# add_sheet is used to create sheet. 
# sheet1 = wb.add_sheet('Sheet 1') 
  
# for i in range(0,len(email_array)):
#     # sheet1.write(i, 0, client_array[i]) 
#     # sheet1.write(i, 1, email_array[i]) 
#     ws.append([client_array[i],company_array[i],email_array[i],business_array[i],doi_array[i],roc_array[i]])


# wb.save('xlwt example.xls') 