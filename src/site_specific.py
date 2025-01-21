# ZiHeR site details for use by ZiherPlus
# Author: Marek Szymański

from selenium.webdriver.common.by import By

Locators = {
    "Login": {
        "MailInput": (By.ID, "user_email"),
        "PasswordInput": (By.ID, "user_password"),
        "LoginButton": (By.CLASS_NAME, "btn-success"),
    },
    "Site": {
        "AccountMenuDropdownLink": (By.XPATH, "//ul[@class='nav navbar-nav pull-right']/li/a"),
        "LogoutLink": (By.XPATH, "//a[@rel='nofollow']"),
        "BankLogLink": (By.LINK_TEXT, "Książka bankowa"),
        "FinLogLink": (By.LINK_TEXT, "Książka finansowa"),
        "FirstFormButton": (By.XPATH, "(//a[@class='btn btn-sm btn-success'])[1]"),
        "SecondFormButton": (By.XPATH, "(//a[@class='btn btn-sm btn-success'])[2]"),
    },
    "Form": {
        "CommitButton": (By.NAME, "commit"),
        "ReturnLink": (By.LINK_TEXT, "Powrót do książki"),
    },
}

FormFieldIDs = {
    'date': 'entry_date',
    'doc_nr': 'entry_document_number',
    'name': 'entry_name',
    'amountFun': lambda idx: f"entry_items_attributes_{idx}_amount",
    'onepFun': lambda idx: f"entry_items_attributes_{idx}_amount_one_percent",
    'grantFun': lambda idx: f"entry_items_attributes_{idx}_item_grants_attributes_0_amount"
}
