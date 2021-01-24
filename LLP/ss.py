import selenium
import PIL
from PIL import Image
from selenium import webdriver


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


driver = webdriver.Chrome(executable_path=r'C:\Users\Ansh\Downloads\chromedriver_win32\chromedriver.exe')
driver.get("http://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do")

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