import os
import pandas as pd
from playwright.sync_api import Page, sync_playwright
import time,warnings
 
warnings.simplefilter("ignore", category=UserWarning)

class Partner_locator:
    def __init__(self):
        self.page = None
        self.Partner_locator_url = "https://freschesolutions.com/partner-locator/"
        self.scraped_data = []
        
    def progress_bar(self, iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
        """
        This function is used to display the progress bar

        Args:
            iteration (int): iteration number
            total (int): total number
            prefix (str): prefix string
            suffix (str): suffix string
            decimals (int): number of decimals
            length (int): length of the progress bar
            fill (str): fill character
            printEnd (str): print end character

        Returns:
            None
        """
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)

        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
        if iteration == total: 
            print(flush=True)
            
    
    def navigate_to_Partner_locator(self, current_page: int = 1):
        """
        This function is used to navigate to Parent locator site

        Args:
            current_page (int): The page number to navigate to. Default is 1.

        Returns:
            True: If the function is successful
            Exception: If the function is unsuccessful
        """
        successfull_navigation = False
        while not successfull_navigation:
            try:
                self.page.goto(f"{self.Partner_locator_url}?page={current_page}")
                retry_count = 0
                while self.get_by_role("heading", name="Partner Locator").all() == []:
                    time.sleep(1)
                    if retry_count == 5:
                        self.page.reload()
                        retry_count = 0
                successfull_navigation = True
            except Exception as e:
                return e.__str__()
        
    def get_navigation_data(self , page: Page):
        """
        This function is used to get the navigation data from the Partner locator site

        Args:
            page (Page): Page object

        Returns:
            partner_type (Locator): Locator for partner type dropdown
            region (Locator): Locator for region dropdown
            prod_solns (Locator): Locator for product solutions dropdown
            partner_type_options (List[Locator]): List of partner type options
            region_options (List[Locator]): List of region options
            prod_soln_options (List[Locator]): List of product solution options
            total_data (int): Total number of data combinations
        """
        self.navigate_to_Partner_locator()
        while True:
            try:
                partner_type = self.page.locator('select[name="partner_type"]')
                region = self.page.locator('select[name="region"]')
                prod_solns = self.page.locator('select[name="prod_soln"]')
                partner_type_options = partner_type.locator('option').all()
                region_options = region.locator('option').all()
                prod_soln_options = prod_solns.locator('option').all()
                total_data = len(partner_type_options) * len(region_options) * len(prod_soln_options)
                if partner_type_options == [] or region_options == [] or prod_soln_options == []:
                    self.page.reload()
                    continue
                else:
                    break
            except Exception as e:
                pass
        return partner_type, region, prod_solns, partner_type_options, region_options, prod_soln_options, total_data
            
    def scrape_Partner_locator_data(self, page: Page):
        """
        This function is used to scrape the Partner_locator site for data
        
        Args:
            page (Page): Page object
            
        Returns:
            scraped_data (dict): Dictionary of scraped data
        """
        while True:
            try:
                partner_type, region, prod_solns, partner_type_options, region_options, prod_soln_options, total_data = self.get_navigation_data(page)
                data_count = 1
                for i in range(1,len(partner_type_options)):
                    
                    for j in range(1,len(region_options)):
                        
                        for k in range(1,len(prod_soln_options)):
                            data_count += 1
                            while True:
                                try:
                                    partner_type.select_option(index=i)
                                    region.select_option(index=j)
                                    prod_solns.select_option(index=k)
                                    break
                                except Exception as e:
                                    pass
                            Search_retry = 0
                            while True:
                                Search_retry += 1
                                try:
                                    page.get_by_role("button", name="Search", exact=True).click()
                                    self.page.wait_for_load_state("load")
                                    break
                                except Exception as e:
                                    if Search_retry == 5:
                                        raise Exception("Search Button Not Found")
                                    pass
                                
                            while True:
                                try:
                                    self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                                    time.sleep(2)
                                    break
                                except Exception as e:
                                    pass
                            pagingItems = []
                            pagingItems_retry = 0
                            while True:
                                pagingItems_retry += 1
                                try:
                                    pagingItems = self.page.locator('//div[@id="pagingItems" ]/ul/li').get_by_role("link").all()
                                    break
                                except Exception as e:
                                    if pagingItems_retry == 5:
                                        raise Exception("Paging Items Not Found")
                                    pass
                            
                            if pagingItems == []:
                                
                                target_data = self.page.locator('[id="body-content"] section[class="section wrap partner-locator"]  [class="product-grid"]  article[class="product-grid--item"]').all()
                                        
                                if target_data != []:
                                    self.scraped_data = self.get_all_data(partner_type_options, region_options, prod_soln_options, data_count, total_data, i, j, k)
                            else:
                                all_href = []
                                for page_item in pagingItems:
                                    
                                    if str(page_item).upper() == "NEXT":
                                        break
                                    else:
                                        href_count = 0
                                        while True:
                                            href_count += 1
                                            try:
                                                all_href.append(page_item.get_attribute("href"))
                                                break
                                            except Exception as e:
                                                if href_count == 5:
                                                    raise Exception("HREF Not Found")
                                                pass
                                for href in all_href:
                                    page_navigation_retry = 0
                                    while True:
                                        page_navigation_retry += 1
                                        try:
                                            self.page.goto(self.Partner_locator_url + href)
                                            self.page.wait_for_load_state("load")
                                            break
                                        except Exception as e:
                                            if page_navigation_retry == 5:
                                                self.close_all_previous_instance()
                                                self.setup()
                                            elif page_navigation_retry == 10:
                                                raise Exception("Page Navigation Failed")
                                            pass
                                    while True:
                                        try:
                                            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                                            time.sleep(2)
                                            break
                                        except Exception as e:
                                            pass
                                    self.scraped_data = self.get_all_data(partner_type_options, region_options, prod_soln_options, data_count, total_data ,i, j, k)
                            self.progress_bar(data_count, total_data, prefix = 'PARTNER LOCATOR:', suffix = 'COMPLETED', length = 100, fill = '█', printEnd = "\r")
                        #break
                break
            except Exception as e:
                pass
        return self.scraped_data
    def get_all_data(self, partner_type_options, region_options, prod_soln_options, data_count, total_data, i=0, j=0, k=0):
        """
        This function is used to get all the data from the Partner locator site

        Args:
            partner_type_options (list): Type of the partner options
            region_options (list): Location of the partner options
            prod_soln_options (list): Product solution options
            data_count (int): current index of the data
            total_data (_type_): Total number of data combinations between partner type, region and product solution
            i (int, optional): partner type index. Defaults to 0.
            j (int, optional): region index. Defaults to 0.
            k (int, optional): product solution index. Defaults to 0.

        Returns:
            scraped_data (list): List of scraped data from the Partner locator site
        """
        
        target_data = self.page.locator('[id="body-content"] section[class="section wrap partner-locator"]  [class="product-grid"]  article[class="product-grid--item"]').all()
        if target_data != []:
            for data in target_data:
                while True:
                    try:
                        header = data.locator('header').inner_text()
                        if header == "":
                            time.sleep(1)
                            continue
                        else:
                            break
                    except Exception as e:
                        pass
                while True: 
                    try:
                        section = data.locator('section').locator('div').all()
                        if section == []:
                            time.sleep(1)
                            continue
                        else:
                            break
                    except Exception as e:
                        pass
                if section != []:
                    while True:
                        section_data = {}
                        section_data["COMPANY NAME"] = header
                        try:
                            section_data["PARTNER TYPE"] = partner_type_options[i].inner_text()
                            section_data["LOCATION"] = region_options[j].inner_text()
                            section_data["PRODUCT SOLUTION"] = prod_soln_options[k].inner_text()
                            break
                        except Exception as e:
                            pass
                    
                    for sec in section:
                        while True:
                            try:
                                key = sec.locator('[class="meta"]').inner_text()
                                if key == "":
                                    time.sleep(1)
                                    continue
                                else:
                                    if key.split(" ")[0].upper() == str(section_data["PARTNER TYPE"]).split(" ")[0].upper():
                                        key = "PHONE NUMBER"
                                    break
                            except Exception as e:
                                pass
                        while True:
                            try:
                                value = sec.locator('//p[@class="meta"]/following-sibling::p').inner_text()
                                if value == "":
                                    time.sleep(1)
                                    continue
                                else:
                                    break
                            except Exception as e:
                                pass
                            
                        section_data[key]= value
                    while True:
                        try:
                            websites = data.locator("footer").locator('a').all()
                            if websites == []:
                                website = ""
                                break
                            else:
                                website = data.locator("footer").locator('a').get_attribute("href")
                                break
                        except Exception as e:
                            pass
                    section_data["WEBSITE"] = website
                    self.scraped_data.append(section_data)                 
        return self.scraped_data
    
    def save_to_excel(self):
        """
        This function is used to save the scraped data to excel
        
        Returns:
            None
        """
        current_path = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_path, "Partner Locator Data")
        os.makedirs(file_path, exist_ok=True)
        file_name = os.path.join(file_path, "Partner_locator.xlsx")
        self.scrape_Partner_locator_data(self.page)
        df = pd.DataFrame(self.scraped_data)
        df.to_excel(file_name, index=False)
        
    def run(self):
        """
        This function is used to run the Partner locator
        
        Returns:
            None
        """
        self.close_all_previous_instance()
        self.setup()
        self.save_to_excel()
        
    def close_all_previous_instance(self):
        """
        This function is used to close all previous instances of the browser ignore any error don't display any error
        
        Returns:
            None
        """
        try:
            os.system("taskkill /f /im chromium.exe /FI \"STATUS eq RUNNING\"")
        except Exception as e:
            pass
        
    def setup(self):
        """
        This function is used to setup the Partner locator

        Returns:
            None
        """
        playwright = sync_playwright().start()
        browser = playwright.chromium.launch_persistent_context(user_data_dir="PARTNER LOCATOR USER DATA", headless=True, args=["--start-maximized"],downloads_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads"),accept_downloads=True)
        page = browser.pages[0]
        self.page = page
        
if __name__ == "__main__":
    
    partner_locator = Partner_locator()
    partner_locator.run()
    