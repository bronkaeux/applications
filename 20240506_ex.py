import pandas as pd
import numpy as np
import re
from datetime import datetime

# loading dictionaries
df = pd.DataFrame(columns=['Менеджер', 'Вид доставки', 'Кол-во', 'Ед.изм.', 'Покупатель',
                           ])

managers = pd.DataFrame({'login': ['ИГО', 'Юра Менеджер', 'Алексей Мельхер'],
                         'manager': ['Игорь Хабаров', 'Юрий', 'Алексей Мельхер']})

units = pd.DataFrame({'data': ['т', 'т.', 'тн', 'тонн', 'кубов'],
                      'unit': ['т', 'т', 'т', 'т', 'куб']})

# entering an application
application = str(input('Введите заявку'))


# processing of application data

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


def delivery_type(text):
    """
    Function for determining the type of delivery
    """
    if "самовывоз" in text.lower():
        return "самовывоз"
    elif "автономка" in text.lower():
        return "автономка доставка"
    else:
        return "доставка"


def find_quantity(text):
    """
    Function for determining the product's quantity
    """
    quantity_match = re.search(r'(?i)кол(\s*-\s*|ичест)?во\s*(?:\D*)(:)?\s*(\d+)', text)
    quantity = int(quantity_match.group(3)) if quantity_match else np.NAN

    return quantity


def find_unit_note(text):
    """
    Function for determining the quantity's unit
    """
    unit_match = re.search(r'(?i)кол(-|ичест)?во\s*(?:\D*)(:)?\s*(\d+)\s*(\D+)\s*\\n\d', text)
    unit = unit_match.group(4).strip() if unit_match else np.NAN

    return unit


def find_purchaser(text):
    """
    Function for determining the purchaser's name
    """
    purchaser_match = re.search(r'(?i)\d+\.\s*(Покупатель\s*(груза)?\s*(:)?)\s*(.+?)\\n', text)
    purchaser = purchaser_match.group(4) if purchaser_match else np.NAN

    return purchaser


def extract_all_data(appl):
    """
    Function to run all functions to eliminate errors in them
    """
    try:

        # assigning a variable with the manager's name
        manager_found = find_manager(appl, managers)

        # assigning a variable with the delivery's type
        delivery_found = delivery_type(appl)

        # assigning a variable with the product's quantity
        quantity_found = find_quantity(appl)

        # assigning a variable with the quantity's unit
        unit_note_found = find_unit_note(appl)

        # finding data in a dictionary
        unit_found = units.loc[units[units['data'] == unit_note_found].index[0], 'unit']

        # assigning a variable with the purchaser's name
        purchaser_found = find_purchaser(appl)



        # creating new df with a string
        new_row = pd.DataFrame({'Менеджер': [manager_found],
                                'Вид доставки': [delivery_found],
                                'Кол-во': [quantity_found],
                                'Ед.изм.': [unit_found],
                                'Покупатель': [purchaser_found],
                                })
    except Exception as e:
        print("Ошибка при извлечении данных:", e)
        new_row = pd.DataFrame({'Менеджер': [appl]})

    return new_row


# assigning a variable with the fun
new_row_df = extract_all_data(application)

# concatenating dfs
df = pd.concat([df, new_row_df], ignore_index=True)

# record updated df
with pd.ExcelWriter("reg.xlsx") as writer:
    df.to_excel(writer, sheet_name="data", index=False)

# print(convert_date, manager_found, delivery_found, product_found, product_notice_found, quantity_found,
#       unit_found, cars_found, organization_found, transshipment_found, purchaser_found, consignee_found,
#       consignee_leg_addr_found, phones_found, unl_addr_found, accept_time_found, note_found)
