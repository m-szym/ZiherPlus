# Custom types used in ZiherPlus
# Author: Marek Szymański
from typing import Literal, NewType

ZihRecord = NewType("ZihRecord", dict[str, str | float])
LogbookType = Literal["bankowa", "finansowa", "inwentarzowa"]
EntryType = Literal["income", "cost"]
RegionType = Literal[
    "dolnyslask", 
    "gornyslasl", 
    "kujawy", 
    "lodz", 
    "lublin", 
    "malopolska",
    "mazowsze",
    "podkarpacie",
    "pomorze",
    "polnocny-zahcod",
    "staropolska",
    "wielkopolska",
]