# Spreadsheet configuration for use by ZiherPlus
# Author: Marek Szymański

from site_specific import FormFieldIDs

'''Dict with details about data arrangement in the Excel spreadsheet

If your spreadsheet is arranged differently reorder the columns by changing 
dictionary keys. Do not change other information here - only rearrange keys 
so that they match your column order

eg. your spreadsheet looks like this
    data | opis ...
    ---------------
    so you will change the key of the item with name 'data' from 2 to 1,
    the key of the item 'opis' from 4 to 2 and so on


'''
COLS = {
    # basic info about the record
    1: {"name": "nr rekordu",                           "type": "misc", "id": "IDX"},
    2: {"name": "data",                                 "type": "misc", "id": FormFieldIDs['date']},
    3: {"name": "nr dokumentu",                         "type": "misc", "id": FormFieldIDs['doc_nr']},
    4: {"name": "opis",                                 "type": "misc", "id": FormFieldIDs['name']},
    # income columns                        
    5: {"name": "składki",                              "type": "income", "id": 0},
    6: {"name": "statutowe",                            "type": "income", "id": 1},
    7: {"name": "dotacje",                              "type": "income", "id": 2},
    8: {"name": "pozostałe",                            "type": "income", "id": 3},
    # aggregates - columns 9, 10, 11
    # costs columns
    12: {"name": "wyposażenie",                         "type": "cost", "id": 0},
    13: {"name": "zużycie materiałów i energii",        "type": "cost", "id": 1},
    14: {"name": "usługi obce",                         "type": "cost", "id": 2},
    15: {"name": "podatki i opłaty",                    "type": "cost", "id": 3},
    16: {"name": "wynagrodzenia",                       "type": "cost", "id": 4},
    17: {"name": "ubezpieczenia i inne świadczenia",    "type": "cost", "id": 5},
    18: {"name": "ubezpieczenia majątkowe",             "type": "cost", "id": 6},
    19: {"name": "podróże służbowe",                    "type": "cost", "id": 7},
    20: {"name": "wyżywienie",                          "type": "cost", "id": 8},
    21: {"name": "nagrody",                             "type": "cost", "id": 9},
    22: {"name": "bilety wstępu",                       "type": "cost", "id": 10},
    23: {"name": "noclegi",                             "type": "cost", "id": 11},
    24: {"name": "transport",                           "type": "cost", "id": 12},
    25: {"name": "inne",                                "type": "cost", "id": 13},
}