import pandas as pd
import numpy as np
import re
from datetime import datetime


# dictionaries downloading
df = pd.read_excel("appl_register.xlsx", sheet_name="data")
managers = pd.read_excel("dictionary.xlsx", sheet_name="managers")
units = pd.read_excel("dictionary.xlsx", sheet_name="units")
unload_addresses = pd.read_excel("dictionary.xlsx", sheet_name="unload_addresses")

appl = input('Введите заявку')

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


def find_manager(text, managers_df):
    """
    Function for getting the manager's name by login from text
    """
    for index, row in managers_df.iterrows():
        login = row['login']
        manager = row['manager']

        if login.lower().replace(' ', '') in text.lower().replace(' ', ''):
            return manager
    return np.NAN


# assigning a variable with the manager's name
manager_found = find_manager(appl, managers)


def delivery_type(text):
    """
    Function for determining the type of delivery
    """
    if "самовывоз" in text.lower():
        return "самовывоз"
    return "доставка"


# assigning a variable with the delivery's type
delivery_found = delivery_type(appl)


def find_product(text):
    """
    Function for determining the type of product
    """
    marca_match = re.search(r'(?i)Марка(:)?\s*(цемента)?\s*(.*?)\n', text)
    marca_value = marca_match.group(3) if marca_match else np.NAN

    return marca_value


# assigning a variable with the product's type
product_found = find_product(appl)


def find_product_notice(text):
    """
    Function for determining the notice for the product
    """
    intermediate_match = re.search(r'(?i)Марка.*?\n(.*?)\n\d', text, re.DOTALL)
    intermediate_value = intermediate_match.group(1) if intermediate_match else np.NAN

    if re.match(r'^[^0-9.].*', intermediate_value):
        return intermediate_value
    else:
        return np.NAN


# assigning a variable with the product's notice
product_notice_found = find_product_notice(appl)


def find_quantity(text):
    """
    Function for determining the product's quantity
    """
    quantity_match = re.search(r'(?i)кол(\s*-\s*|ичест)?во\s*(:)?\s*(\d+)', text)
    quantity = int(quantity_match.group(3)) if quantity_match else np.NAN

    return quantity


# assigning a variable with the product's quantity
quantity_found = find_quantity(appl)


def find_unit_note(text):
    """
    Function for determining the quantity's unit
    """
    unit_match = re.search(r'(?i)кол(-|ичест)?во\s*(:)?\s*(\d+)\s*(\D+)\s*', text)
    unit = unit_match.group(4).strip() if unit_match else np.NAN

    return unit


# assigning a variable with the quantity's unit
unit_note_found = find_unit_note(appl)

# finding data in a dictionary
unit_found = units.loc[units[units['data'] == unit_note_found].index[0], 'unit']


def find_car(text):
    """
    Function for determining the car's numbers
    """
    car_match = re.findall(r'(?i)[А-Я]{1}\d{3}[А-Я]{2}\d{3}\s*\w*', text)
    cars = car_match if car_match else np.NAN

    return cars


# assigning a variable with car numbers and writing them without list brackets
cars_found = ', '.join(find_car(appl)) if find_car(appl) is not np.NAN else np.NAN


def find_organization(text):
    """
    Function for determining our organization's name
    """
    organization_match = re.search(r'(?i)(продажа\s+от:|(продажа)?\s*от\s+(клиента)?\s*(:)?)\s*(.+?)\n', text)
    organization = organization_match.group(5) if organization_match else np.NAN

    return organization


# assigning a variable with our organization's name
organization_found = find_organization(appl)


def find_transshipment(text):
    """
    Function for determining the transshipment's point
    """
    transshipment_match = \
        re.search(r'(?i)\d+\.\s*(С\s+(перевалки)?\s*|Завод\s*(отгрузки)?\s*(:)?|Перевалка\s*(:)?)\s*(.+?)\n',
                  text)
    transshipment = transshipment_match.group(6) if transshipment_match else np.NAN

    return transshipment


# assigning a variable with the transshipment's point
transshipment_found = find_transshipment(appl)



