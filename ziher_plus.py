# Driver for automating the importing of Excel data into ZiHeR platform
# Author: Marek Szymański
# TODO: add suport for 1% and ROHiS grants

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from typing import Optional
try:
    from typing import Self  
except ImportError:
    from typing_extensions import Self  

from zih_loader import iter_data, load_workbook, load_worksheet, print_record
from site_specific import Locators
from zih_types import ZihRecord, LogbookType, EntryType


class ZiherPlus:
    '''Automation driver for ZiHeR
    
    :param driver: selenium.WebDriver to be used, configure as needed, but will override implicit waits
    :param human_control: will the procces be controlled by human or fully automated
    :param filename: use to immediately load workbook
    :param sheetname: use to immediately load worksheet
    
    :ivar driver: selenium.WebDriver used to operate the browser
    :ivar __region: ZiHeR regional subdomain
    :ivar __human_control: is the driver human-supervised
    :ivar __workbook: Excel workbook - the data source
    :ivar __worksheet: specific worksheet of the __workbook 

    
    '''

    # ==========================================================
    # API
    # ==========================================================

    def __init__(
        self,
        driver: WebDriver,
        human_control: bool = True,
        filename: Optional[str] = None,
        sheetname: Optional[str] = None,
    ):
        self._driver = driver
        self.__human_control = human_control
        self.__workbook = filename
        self.__worksheet = sheetname

        self._driver.implicitly_wait(15)
        if filename:
            self.__workbook, self.__worksheet = load_workbook(filename, sheetname)

    def quit(self) -> None:
        '''Closes ZiherPlus and the controlled browser
        
        Should also mean it'll log out of ZiHeR, but the behaviour is controlled by the app,
        and not ZiherPlus.
        '''
        self._driver.quit()

    def login(self, email: str, password: str, region: str = "pomorze") -> Self:
        """Log into ziher using the provided credentials

        Can be run from anywhere, because it uses direct GET.
        If succesful leaves the driver on the welcome page, signed in.

        :param email: email used for signing in
        :param password: password for signing in
        :param region: regional subdomain to use
        """
        self._driver.get(f"https://ziher.zhr.pl/{region}/users/sign_in")

        self._driver.find_element(*Locators["Login"]["MailInput"]).send_keys(email)
        self._driver.find_element(*Locators["Login"]["PasswordInput"]).send_keys(
            password
        )
        self._driver.find_element(*Locators["Login"]["LoginButton"]).click()

        return self

    def logout(self) -> Self:
        """Log out of ziher

        Can be run from anywhere on ziher plaftorm, as long as User-Account submenu is available.
        Leaves the driver on the ziher login page.
        Ends method chaining.
        """
        # self.__driver.find_element(*Locators["Site"]["AccountMenuDropdownLink"]).click()
        # self.__driver.find_element(*Locators["Site"]["LogoutLink"]).click()

        self.__use_dropdown(dropdown_text="@zhr.pl", option_text="Wyloguj się")

    def load(self, filename: str, sheetname: Optional[str] = None) -> Self:
        """Loads the Excel file: `filename` and optionally opens worksheet `sheetname` or the active one

        :param filename: path to Excel file
        :param sheetname: worksheet to open, leave out to open the active one
        """
        if self.__workbook:
            self.__workbook.close()
        self.__workbook, self.__worksheet = load_workbook(filename, sheetname)
        return self

    def worksheet(self, sheetname: Optional[str] = None) -> Self:
        """Loads the worksheet specified by `sheetname` of the current workbook or the active one

        :param sheetname: name of the worksheet to load, leave out to open the active one
        """
        self.__worksheet = load_worksheet(self.__workbook, sheetname)
        return self

    def send(
        self,
        logbook: LogbookType,
        min_row: int,
        max_row: int,
        sheetname: Optional[str] = None,
        filename: Optional[str] = None
    ) -> Self:
        '''Import Excel data into ziher

        :param logbook: string identifying targeted logbook
        :param min_row: nr of the first row of data to be imported
        :param max_row: nr of the last row of data to be imported
        :param sheetname: use to change worksheet
        :param filename: use to change Excel file
        '''

        if filename:
            self.load(filename, sheetname)
        elif sheetname:
            self.worksheet(sheetname)

        self.__open_log(logbook)

        for i, record in enumerate(iter_data(self.__worksheet, min_row, max_row)):
            if self.__human_control:
                print_record(record, i)
            else:
                print(f"Record {i} - from Excel {record['IDX']}")

            try:
                self.__fill_form(record)
            except Exception as err:
                print(f"Error: {type(err)}")
                if self.__human_control:
                    print(err)
            
            time.sleep(2)

        return self

    # Convenience constructors for different browsers

    @classmethod
    def Firefox(cls, human_control: bool = True):
        '''ZiherPlus driver for Firefox'''
        return cls(webdriver.Firefox(), human_control)

    @classmethod
    def Chrome(cls, human_control: bool = True):
        '''ZiherPlus driver for Chrome'''
        return cls(webdriver.Chrome(), human_control)

    @classmethod
    def Edge(cls, human_control: bool = True):
        '''ZiherPlus driver for MSEdge'''
        return cls(webdriver.Edge(), human_control)

    @classmethod
    def Safari(cls, human_control: bool = True):
        '''ZiherPlus driver for Safari'''
        return cls(webdriver.Safari(), human_control)
        


    # ==========================================================
    # Private methods
    # ==========================================================

    def __open_log(self, book: LogbookType):
        """Switch to another logbook

        :param book: string identifying logbook
        """
        if book == "bankowa":
            self._driver.find_element(*Locators["Site"]["BankLogLink"]).click()
        elif book == "finansowa":
            self._driver.find_element(*Locators["Site"]["FinLogLink"]).click()
        elif book == "inwentarzowa":
            self._driver.find_element(*Locators["Site"]["InventoryLogLink"]).click()

    def __open_form(self, formtype: EntryType):
        """Opens correct form - either for declaring income or cost

        :param formtype: which form to open
        """
        if formtype == "income":
            self._driver.find_element(*Locators['Site']['FirstFormButton']).click()
        elif formtype == "cost":
            self._driver.find_element(*Locators['Site']['SecondFormButton']).click()

    def __fill_form(self, entry_data: ZihRecord):
        """Opens, fills and commits new form with provided single `entry_data`

        :param entry_data: dict conatining input_ids, respective input values
                           and other info that's not sent to the form
        """
        self.__open_form(entry_data["type"])

        for k, v in entry_data.items():
            if k in ["type", "category", "IDX"]:
                continue
            self._driver.find_element(By.ID, k).send_keys(v)

        if self.__human_control:
            self._human_commit()
        else:
            self._driver.find_element(*Locators['Form']['CommitButton']).click()

        return self

    def _human_commit(self) -> None:
        """Asks and waits for human approval or lack thereof before commiting new entry"""
        commit = input("Commit ?: [y/n] ")
        if commit == "y":
            self._driver.find_element(*Locators['Form']['CommitButton']).click()
        elif commit == "n":
            self._driver.find_element(*Locators["Form"]["ReturnLink"]).click()

    def __use_dropdown(self, dropdown_text: str, option_text: str):
        """Opens dropdown identified by `dropdown_text` and clicks the options with text containing `option_text`

        :param dropdown_text: text to identify correct dropdown, can be only a fragment of the full link text, but must be explicit to this dropdown
        :param option_text: text to identify correct dropdown option, can be only a fragment of the full link text, but must be explicit to this option
        """
        dropdown_node_xpth = "//*[contains(@class, 'dropdown')]"
        self._driver.find_element(
            By.XPATH, dropdown_node_xpth + f"/a[contains(text(), '{dropdown_text}')]"
        ).click()  # open the dropdown
        time.sleep(1)
        self._driver.find_element(
            By.XPATH, dropdown_node_xpth + f"/li/a[contains(text(), '{option_text}')]"
        ).click()  # click specified option

        return self

    # FIXME
    def __click_button(self, button_text: str):
        """Clicks button that contains `button_text`

        :param button_text: part or whole of the button text
        """
        self._driver.find_element(
            By.XPATH,
            f"//a[contains(@class, 'btn') and contains(text(), '{button_text}')]",
        ).click()
        return self


class ZiherPlusSafeMode(ZiherPlus):
    '''Safe version of ZiherPlus driver - this one will NEVER commit any entry. 
    Use to get a preview without importing unwanted data.

    Human-control is ON by default.
    After asking for commit confirmation this driver will ALWAYS return without commiting.
    '''
    def __init__(self, driver: WebDriver):
        super().__init__(driver=driver, human_control=True)

    # Convenience constructors for different browsers
    
    @classmethod
    def Firefox(cls):
        '''ZiherPlusSafeMode driver for Firefox'''
        return cls(webdriver.Firefox())
    
    @classmethod
    def Chrome(cls):
        '''ZiherPlusSafeMode driver for Chrome'''
        return cls(webdriver.Chrome())
    
    @classmethod
    def Edge(cls):
        '''ZiherPlusSafeMode driver for MSEdge'''
        return cls(webdriver.Edge())
    
    @classmethod
    def Safari(cls):
        '''ZiherPlusSafeMode driver for Safari'''
        return cls(webdriver.Safari())

    # ==========================================================
    # Private methods
    # ==========================================================
    
    def _human_commit(self) -> None:
        """Asks and waits for human input before returning,
            NEVER commits entry
        """
        input("Commit ?: [n/n] ")
        self._driver.find_element(*Locators["Form"]["ReturnLink"]).click()