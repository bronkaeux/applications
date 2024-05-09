import pandas as pd
import numpy as np
import re
import traceback
from datetime import datetime
import functools


# Error Handling Decorator
def handle_errors(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_message = f"Error in {func.__name__ if hasattr(func, '__name__') else 'unknown method'}: {e}"
            args[0].error_log = args[0].error_log.append({'Ошибка': error_message, 'Заявка': args[1]},
                                                         ignore_index=True)
            return np.nan

    return wrapper


class ApplicationProcessor:
    def __init__(self, applications_df, managers_df, units_df, unload_addresses_df):
        self.applications_df = applications_df
        self.managers_df = managers_df
        self.units_df = units_df
        self.unload_addresses_df = unload_addresses_df
        self.error_log = pd.DataFrame(columns=['Ошибка', 'Заявка'])

    @handle_errors
    def find_manager(self, text):
        """
        Method for getting the manager's name by login from text
        """
        for index, row in self.managers_df.iterrows():
            login = row['login']
            manager = row['manager']

            if login.lower().replace(' ', '') in text.lower().replace(' ', ''):
                return manager
        return np.nan

    @staticmethod
    @handle_errors
    def find_dates(text):
        """
        Method for dates extracting
        """
        date_pattern = r'\d{2}\.(?:0[1-9]|1[0-2])(?:\.\d{2}(?:\d{2})?)?'
        dates = re.findall(date_pattern, text)
        return dates

    @staticmethod
    @handle_errors
    def convert_to_full_year(date_str):
        """
        Method for formatting the date into the specified format
        """
        current_year = datetime.now().year
        parts = date_str.split('.')

        if len(parts) == 2:
            date_str += f".{current_year}"
        else:
            date_str = parts[0] + '.' + parts[1] + f".{current_year}"
        datetime_obj = datetime.strptime(date_str, '%d.%m.%Y')

        return datetime_obj.strftime('%d.%m.%Y')

    @staticmethod
    @handle_errors
    def delivery_type(text):
        """
        Method for determining the type of delivery
        """
        if "самовывоз" in text.lower():
            return "самовывоз"
        elif "автономка" in text.lower():
            return "автономка доставка"
        else:
            return "доставка"

    @staticmethod
    @handle_errors
    def find_product(text):
        """
        Method for determining the type of product
        """
        marca_match = re.search(r'(?i)Марка(:)?\s*(цемента)?\s*(.*?)\\n', text)
        marca_value = marca_match.group(3) if marca_match else np.nan
        return marca_value

    @staticmethod
    @handle_errors
    def find_product_notice(text):
        """
        Method for determining the notice for the product
        """
        intermediate_match = re.search(r'(?i)Марка.*?\\n(.*?)\\n\d', text, re.DOTALL)
        intermediate_value = intermediate_match.group(1) if intermediate_match else np.nan

        if re.match(r'^[^0-9.].*', intermediate_value):
            return intermediate_value
        else:
            return np.nan

    @staticmethod
    @handle_errors
    def find_quantity(text):
        """
        Method for determining the product's quantity
        """
        quantity_match = re.search(r'(?i)кол(\s*-\s*|ичест)?во\s*(?:\D*)(:)?\s*(\d+)', text)
        quantity = int(quantity_match.group(3)) if quantity_match else np.nan
        return quantity

    @handle_errors
    def find_unit_note(self, text):
        """
        Method for determining the quantity's unit
        """
        unit_match = re.search(r'(?i)кол(-|ичест)?во\s*(?:\D*)(:)?\s*(\d+)\s*(\D+)\s*\\n\d', text)
        unit = unit_match.group(4).strip() if unit_match else np.nan

        if unit:
            unit_index = self.units_df[self.units_df['data'] == unit].index
            if not unit_index.empty:
                unit_found = self.units_df.loc[unit_index[0], 'unit']
                return unit_found
        return np.nan

    @staticmethod
    @handle_errors
    def find_car(text):
        """
        Method for determining the car's numbers
        """
        car_match = re.findall(r'(?i)[А-Я]{1}\d{3}[А-Я]{2}\d{3}\s*\w*', text)
        cars = car_match if car_match else np.nan
        return cars

    @staticmethod
    @handle_errors
    def find_organization(text):
        """
        Method for determining our organization's name
        """
        organization_match = re.search(r'(?i)(продажа\s+от:|(продажа)?\s*от\s+(клиента)?\s*(:)?)\s*(.+?)\\n', text)
        organization = organization_match.group(5) if organization_match else np.nan
        return organization

    @staticmethod
    @handle_errors
    def find_transshipment(text):
        """
        Method for determining the transshipment's point
        """
        transshipment_match = \
            re.search(r'(?i)\d+\.\s*(С\s+(перевалки)?\s*|Завод\s*(отгрузки)?\s*(:)?|Перевалка\s*(:)?)\s*(.+?)\\n',
                      text)
        transshipment = transshipment_match.group(6) if transshipment_match else np.nan
        return transshipment

    @staticmethod
    @handle_errors
    def find_purchaser(text):
        """
        Method for determining the purchaser's name
        """
        purchaser_match = re.search(r'(?i)\d+\.\s*(Покупатель\s*(груза)?\s*(:)?)\s*(.+?)\\n', text)
        purchaser = purchaser_match.group(4) if purchaser_match else np.nan
        return purchaser

    @staticmethod
    @handle_errors
    def find_consignee(text):
        """
        Method for determining the consignee's name
        """
        consignee_match = re.search(
            r'(?i)\d+\.\s*(?:Грузопол\w*\s*(?::)?|Грузопол\w*\s*\(при\s*оформ\w*\s*ттн\)\s*(?::)?)\s*(.+?)\\n',
            text)
        consignee = consignee_match.group(1).split(':')[-1].strip() if consignee_match else np.nan
        return consignee

    @staticmethod
    @handle_errors
    def find_consignee_leg_addr(text):
        """
        Method for determining the legal consignee's address
        """
        find_consignee_leg_addr_match = re.search(
            r'(?i)\d+\.\s*(юр\w*\s*(?:\.)?\s*адрес\s*грузополучателя|адрес\s*грузополучателя\s*\(юр\w*\s*(?:\.)?\))\s*(?::)?\s*(.+?)\\n',
            text)
        find_consignee_leg_addr = find_consignee_leg_addr_match.group(2) if find_consignee_leg_addr_match else np.nan
        return find_consignee_leg_addr

    @handle_errors
    def find_unl_addr(self, text):
        """
        Method for determining unload addresses
        """
        for index, row in self.unload_addresses_df.iterrows():
            pattern1 = row['name']
            pattern2 = row['place']
            address = row['address']

            if pattern1.lower().replace(' ', '') and pattern2.lower().replace(' ', '') \
                    in text.lower().replace(' ', ''):
                return address
        return np.nan

    @staticmethod
    @handle_errors
    def find_phones(text):
        """
        Method for determining the phone numbers
        """
        phone_numbers = re.findall(r'\+?\d{1,3}[\s-]?\(?\d{3}\)?[\s-]?\d{2,3}[\s-]?\d{2}[\s-]?\d{2}', text)
        phone_numbers = phone_numbers if phone_numbers else np.nan
        return phone_numbers

    @staticmethod
    @handle_errors
    def find_time(text):
        """
        Method for determining the unloading time
        """
        time_match = re.search(r'(?i)(время)?\s*при(ё|е)мк(и|а)(?::)?\s*(.*?)\s*\\n', text)
        time = time_match.group(4) if time_match else np.nan
        return time

    @staticmethod
    @handle_errors
    def find_note(text):
        """
        Method for determining the note
        """
        if 'оплата' in text.lower():
            note_match = re.search(r'(?i)(оплата)\s*(?::)?\s*(.*?)\\n', text)
            note = '{} {}'.format(note_match.group(1).strip(), note_match.group(2).strip() if note_match else np.nan)
            return note
        else:
            return np.nan

    def process_application(self, application):
        """
        Method to process an application and update the dataframe
        """
        try:
            manager_found = self.find_manager(application)
            found_dates = ApplicationProcessor.find_dates(application)
            if found_dates:
                formatted_date = self.convert_to_full_year(found_dates[0])
            else:
                formatted_date = np.nan
            delivery_found = ApplicationProcessor.delivery_type(application)
            product_found = ApplicationProcessor.find_product(application)
            product_notice_found = ApplicationProcessor.find_product_notice(application)
            quantity_found = ApplicationProcessor.find_quantity(application)
            unit_found = self.find_unit_note(application)
            num_str = round(quantity_found / 35 if (unit_found == 'т' or 'цем' in product_found.lower())
                            else quantity_found / 40)
            cars_found = ', '.join(ApplicationProcessor.find_car(application)) if \
                ApplicationProcessor.find_car(application) is not np.nan else np.nan
            organization_found = ApplicationProcessor.find_organization(application)
            transshipment_found = ApplicationProcessor.find_transshipment(application)
            purchaser_found = ApplicationProcessor.find_purchaser(application)
            consignee_found = ApplicationProcessor.find_consignee(application)
            consignee_leg_addr_found = ApplicationProcessor.find_consignee_leg_addr(application)
            unl_addr_found = self.find_unl_addr(application)
            ApplicationProcessor.find_phones(application)
            phones_found = ', '.join(ApplicationProcessor.find_phones(application)) if \
                ApplicationProcessor.find_phones(application) is not np.nan else np.nan
            accept_time_found = ApplicationProcessor.find_time(application)
            note_found = ApplicationProcessor.find_note(application)

            new_row = pd.DataFrame({'Менеджер': [manager_found],
                                    'Дата': [formatted_date],
                                    'Вид доставки': [delivery_found],
                                    'Товар': [product_found],
                                    'Примечание к Товару': [product_notice_found],
                                    'Кол-во': [35 if (unit_found == 'т' or 'цем' in product_found.lower()) else 40],
                                    'Ед.изм.': [unit_found],
                                    'Машина/Водитель': [cars_found],
                                    'Продавец': [organization_found],
                                    'Откуда': [transshipment_found],
                                    'Покупатель': [purchaser_found],
                                    'Грузополучатель': [consignee_found],
                                    'Юр. адрес грузополучателя': [consignee_leg_addr_found],
                                    'Адрес пункта разгрузки': [unl_addr_found],
                                    'Контакт гп': [phones_found],
                                    'Время приемки': [accept_time_found],
                                    'Примечание Иное': [note_found],
                                    'Текст заявки': [application.replace('\\n', ' ')]})

            new_rows = pd.concat([new_row] * num_str, ignore_index=True)
            self.applications_df = pd.concat([self.applications_df, new_rows], ignore_index=True)
            self.applications_df['№ Заявки'] = self.applications_df.index + 1

        except Exception as e:
            traceback_str = traceback.format_exc()
            self.error_log = pd.concat([self.error_log, pd.DataFrame({'Ошибка': [traceback_str],
                                                                      'Заявка': [application]})],
                                       ignore_index=True)

            new_row = pd.DataFrame({'Текст заявки': [application]})
            self.applications_df = pd.concat([self.applications_df, new_row], ignore_index=True)
            self.applications_df['№ Заявки'] = self.applications_df.index + 1

    def save_to_excel(self, filename):
        """
        Method to save the dataframe and error log to Excel
        """

        with pd.ExcelWriter(filename) as writer:
            self.applications_df.to_excel(writer, sheet_name="data", index=False)
            self.error_log.to_excel(writer, sheet_name="ошибки", index=False)


# loading dictionaries
filename = "appl_register.xlsx"
applications = pd.read_excel(filename, sheet_name="data")
managers = pd.read_excel("dictionary.xlsx", sheet_name="managers")
units = pd.read_excel("dictionary.xlsx", sheet_name="units")
unload_addresses = pd.read_excel("dictionary.xlsx", sheet_name="unload_addresses")

# entering an application
application = str(input('Введите заявку: '))

# Create an instance of ApplicationProcessor
processor = ApplicationProcessor(applications, managers, units, unload_addresses)

# Process the application
processor.process_application(application)

# Save to Excel
processor.save_to_excel(filename)
