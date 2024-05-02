import pandas as pd
import numpy as np
import re
from datetime import datetime


# dictionaries downloading
df = pd.read_excel("appl_register.xlsx", sheet_name="data")
managers = pd.read_excel("dictionary.xlsx", sheet_name="managers")
units = pd.read_excel("dictionary.xlsx", sheet_name="units")
unload_addresses = pd.read_excel("dictionary.xlsx", sheet_name="unload_addresses")

appl = str(input('Введите заявку'))


# processing of application data
# extracting dates from application
date_pattern = r'\b\d{2}\.(?:0[1-9]|1[0-2])(?:\.\d{2}(?:\d{2})?)?\b'
dates = re.findall(date_pattern, appl)


def convert_to_full_year(date_str):
    """
    Function for formatting the date into the specified format
    """
    current_year = datetime.now().year
    parts = date_str.split('.')

    try:
        if len(parts) == 2:
            date_str += f".{current_year}"

        datetime_obj = datetime.strptime(date_str, '%d.%m.%Y')
    except ValueError:
        try:
            datetime_obj = datetime.strptime(date_str, '%d.%m.%y')
        except ValueError:

            raise ValueError("Неподдерживаемый формат даты")

    return datetime_obj.strftime('%d.%m.%Y')


# assigning a variable with a date
convert_date = convert_to_full_year(dates[0])

print(convert_date)
