from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import pandas as pd
import os


software_componenets_list = []
support_package_list = []

user_agent = "enter your own user agent information here"

reading_excel_file = 'Write the excel file path where the links will be read'
writing_excel_file = 'Write the excel file path where the scraping output will be saved'

enter_username = "enter your username"
enter_password = "enter your password"

chrome_options = Options()
chrome_options.add_argument(f"user-agent={user_agent}")


class SapScraping():
    # read input excel file and go to website
    def __init__(self):        
        links = self.reading_excel(reading_excel_file)

        self.driver = Chrome(options=chrome_options)
        self.driver.get(links[0])


    # log in to the website
    def login(self):
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="j_username"]'))
        )
        user_name = self.driver.find_element(By.XPATH, '//*[@id="j_username"]')
        user_name.send_keys(enter_username)

        continue_button = self.driver.find_element(By.XPATH, '//*[@id="logOnFormSubmit"]')
        continue_button.click()
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="truste-consent-button"]'))
        )

        cookie_button = self.driver.find_element(By.XPATH, '//*[@id="truste-consent-button"]')
        cookie_button.click()
        sleep(5)
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))
        )

        password = self.driver.find_element(By.XPATH, '//*[@id="password"]')
        password.send_keys(enter_password)

        signin_button = self.driver.find_element(By.XPATH, '//*[@id="root"]/main/div[1]/div/div/div[2]/div[2]/div/div/div/div/form/div/div[2]/button')
        signin_button.click()
        WebDriverWait(self.driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="__xmlview15--objectPageLayout_NoteKBADisplay-anchBar-__section7-anchor"]/div[1]'))
        )


    # scrape the necessary data
    def scraping(self, link):
        self.driver.get(link)
        sleep(10)

        software_componenets_base = self.driver.find_element(By.XPATH, '//*[@id="__xmlview15--idTabSoftCom-tblBody"]')
        software_componenets = software_componenets_base.find_elements(By.TAG_NAME, 'tr')

        for i, sc in enumerate(software_componenets):
            software_ID = self.driver.find_element(By.XPATH, '//*[@id="__xmlview15--idObjectPageHeader-innerTitle"]').text.split(" ")[0]
            software = sc.find_element(By.XPATH, '//*[@id="__text34-__xmlview15--idTabSoftCom-' + str(i) + '"]').text
            software_from = sc.find_element(By.XPATH, '//*[@id="__text35-__xmlview15--idTabSoftCom-' + str(i) + '"]').text
            software_to = sc.find_element(By.XPATH, '//*[@id="__text36-__xmlview15--idTabSoftCom-' + str(i) + '"]').text

            software_componenets_list.append({'ID': software_ID, 'Software Componenets' : software, 'From' : software_from, 'To' : software_to})


        support_packages_base = self.driver.find_element(By.XPATH, '//*[@id="__xmlview15--idTabSP-tblBody"]')
        support_packages = support_packages_base.find_elements(By.TAG_NAME, 'tr')

        for i, sp in enumerate(support_packages):
            support_ID = self.driver.find_element(By.XPATH, '//*[@id="__xmlview15--idObjectPageHeader-innerTitle"]').text.split(" ")[0]
            sc_version = sp.find_element(By.XPATH, '//*[@id="__item24-__clone' + str(i) + '-cell0"]').text
            support_package = sp.find_element(By.XPATH, '//*[@id="__link27-__clone' + str(i) + '"]').text
            
            support_package_list.append({'ID' : support_ID, 'Software Component Version' : sc_version, 'Support Package' : support_package})

        
    # write scraped data to output excel file   
    def writing_excel(self, file_path):
        df_software = pd.DataFrame(software_componenets_list)
        df_support = pd.DataFrame(support_package_list)

        if not os.path.exists(file_path):

            with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:
                df_software.to_excel(writer, sheet_name="Software Componenets", index=False)
                df_support.to_excel(writer, sheet_name="Support Package", index=False)

        else:

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                existing_data_software = pd.read_excel(file_path, sheet_name="Software Componenets")
                existing_data_support = pd.read_excel(file_path, sheet_name="Support Package")
                
                combined_data_software = pd.concat([existing_data_software, df_software], ignore_index=True)
                combined_data_support = pd.concat([existing_data_support, df_support], ignore_index=True)
                
                combined_data_software.to_excel(writer, sheet_name="Software Componenets", index=False)
                combined_data_support.to_excel(writer, sheet_name="Support Package", index=False)


    # read input excel file
    def reading_excel(self, file_path):
        df_excel = pd.read_excel(file_path)
        link_column = df_excel["Link"].dropna().tolist()

        return link_column


if __name__ == '__main__':
    sap_scraping = SapScraping()
    sap_scraping.login()
    links = sap_scraping.reading_excel(reading_excel_file)

    for link in links:
        sap_scraping.scraping(link)
        sap_scraping.writing_excel(writing_excel_file)
            
            
    sap_scraping.driver.quit()
