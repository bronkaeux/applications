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
dates = dates[0]


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
convert_date = convert_to_full_year(dates)


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


def find_purchaser(text):
    """
    Function for determining the purchaser's name
    """
    purchaser_match = \
        re.search(r'(?i)\d+\.\s*(Покупатель\s*(груза)?\s*(:)?)\s*(.+?)\n',
                  text)
    purchaser = purchaser_match.group(4) if purchaser_match else np.NAN

    return purchaser


# assigning a variable with the purchaser's name
purchaser_found = find_purchaser(appl)


def find_consignee(text):
    """
    Function for determining the consignee's name
    """
    consignee_match = \
        consignee_match = re.search(
        r'(?i)\d+\.\s*(?:Грузополучатель\s*(?::)?|Грузополучатель\s*\(при\s*оформлении\s*ттн\)\s*(?::)?)\s*(.+?)\n',
        text)
    consignee = consignee_match.group(1).split(':')[-1].strip() if consignee_match else np.NAN

    return consignee


# assigning a variable with the consignee's name
consignee_found = find_consignee(appl)


def find_consignee_leg_addr(text):
    """
    Function for determining the legal consignee's address
    """
    find_consignee_leg_addr_match = \
        re.search(r'(?i)\d+\.\s*(Юр\s*\.)?\s*(Адрес\s*грузополучателя)\s*(\(юрид\))?\s*(:)?\s*(.+?)\n',
                  text)
    find_consignee_leg_addr = find_consignee_leg_addr_match.group(5) if find_consignee_leg_addr_match else np.NAN

    return find_consignee_leg_addr


# assigning a variable with legal consignee's address
consignee_leg_addr_found = find_consignee_leg_addr(appl)


def find_phones(text):
    """
    Function for determining the phone numbers
    """
    phone_numbers = re.findall(r'\+?\d{1,3}[\s-]?\(?\d{3}\)?[\s-]?\d{2,3}[\s-]?\d{2}[\s-]?\d{2}(?=\s|$)', text)
    phone_numbers = phone_numbers if phone_numbers else np.NAN

    return phone_numbers


# assigning a variable with phone numbers and writing them without list brackets
phones_found = ', '.join(find_phones(appl)) if find_phones(appl) is not np.NAN else np.NAN


def find_unl_addr(text, unload_addresses_df):
    """
    Function for determining unload addresses
    """
    for index, row in unload_addresses_df.iterrows():
        pattern1 = row['pattern1']
        pattern2 = row['pattern2']
        address = row['address']

        if pattern1.lower().replace(' ', '') and pattern2.lower().replace(' ', '') \
                in text.lower().replace(' ', ''):
            return address
    return np.NAN


# assigning a variable with unload addresses
unl_addr_found = find_unl_addr(appl, unload_addresses)


def find_time(text):
    """
    Function for determining the unloading time
    """
    time_match = re.search(r'(?i)(время)?\s*при(ё|е)мк(и|а)(?::)?\s*(.*)', text)
    time = time_match.group(4) if time_match else np.NAN

    return time


# assigning a variable with unloading time
accept_time_found = find_time(appl)


def find_note(text):
    """
    Function for determining the note
    """
    note_match = re.search(r'(?i)(.*)\s*(?::)?\s*(оплата)\s*(?::)?\s*(.*)', text)
    note = '{} {} {}'.format(note_match.group(1).strip(), note_match.group(2).strip(),
                             note_match.group(3).strip()) if note_match else np.NAN

    return note


# assigning a variable with the note
note_found = find_note(appl)


# creating new df with a string
new_row_df = pd.DataFrame({'Менеджер': [manager_found],
                           'Вид доставки': [delivery_found],
                           'Дата': [convert_date],
                           'Товар': [product_found],
                           'Примечание к Товару': [product_notice_found],
                           'Кол-во': [quantity_found],
                           'Ед.изм.': [unit_found],
                           'Машина/Водитель': [cars_found],
                           'Продавец': [organization_found],
                           'Откуда': [transshipment_found],
                           'Покупатель': [purchaser_found],
                           'Грузополучатель': [consignee_found],
                           'Адрес грузополучателя (юрид)': [consignee_leg_addr_found],
                           'Адрес пункта разгрузки': [unl_addr_found],
                           'Контакт гп': [phones_found],
                           'Время приемки': [accept_time_found],
                           'Примечание Иное': [note_found],
                          })

# concatenating dfs
df = pd.concat([df, new_row_df], ignore_index=True)

# record updated df
with pd.ExcelWriter("appl_register.xlsx") as writer:
    df.to_excel(writer, sheet_name="data", index=False)

