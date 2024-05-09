import pandas as pd
import numpy as np
import re
import traceback
from datetime import datetime


class ApplicationProcessor:
    def __init__(self, applications_df, managers_df, units_df, unload_addresses_df):
        self.applications_df = applications_df
        self.managers_df = managers_df
        self.units_df = units_df
        self.unload_addresses_df = unload_addresses_df
        self.error_log = pd.DataFrame(columns=['Ошибка', 'Заявка'])

    def find_manager(self, text):
        """
        Method for getting the manager's name by login from text
        """
        try:
            for index, row in self.managers_df.iterrows():
                login = row['login']
                manager = row['manager']

                if login.lower().replace(' ', '') in text.lower().replace(' ', ''):
                    return manager
            return np.nan

        except Exception as e:
            self.error_log.append(f"Error in the find_manager method: {e}")
            return np.nan

    @staticmethod
    def find_dates(text):
        """
        Method for dates extracting
        """
        try:
            date_pattern = r'\d{2}\.(?:0[1-9]|1[0-2])(?:\.\d{2}(?:\d{2})?)?'
            dates = re.findall(date_pattern, text)
            return dates

        except Exception as e:
            self.error_log.append(f"Error in the find_dates method: {e}")
            return np.nan

    @staticmethod
    def convert_to_full_year(date_str):
        """
        Method for formatting the date into the specified format
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
                raise ValueError("Unsupported date format")

        return datetime_obj.strftime('%d.%m.%Y')

    @staticmethod
    def delivery_type(text):
        """
        Method for determining the type of delivery
        """
        try:
            if "самовывоз" in text.lower():
                return "самовывоз"
            elif "автономка" in text.lower():
                return "автономка доставка"
            else:
                return "доставка"

        except Exception as e:
            self.error_log.append(f"Error in the delivery_type method: {e}")
            return np.nan

    @staticmethod
    def find_product(text):
        """
        Function for determining the type of product
        """
        try:
            marca_match = re.search(r'(?i)Марка(:)?\s*(цемента)?\s*(.*?)\\n', text)
            marca_value = marca_match.group(3) if marca_match else np.NAN
            return marca_value

        except Exception as e:
            self.error_log.append(f"Error in the find_product method: {e}")
            return np.nan

    @staticmethod
    def find_product_notice(text):
        """
        Function for determining the notice for the product
        """
        try:
            intermediate_match = re.search(r'(?i)Марка.*?\\n(.*?)\\n\d', text, re.DOTALL)
            intermediate_value = intermediate_match.group(1) if intermediate_match else np.NAN

            if re.match(r'^[^0-9.].*', intermediate_value):
                return intermediate_value
            else:
                return np.nan

        except Exception as e:
            self.error_log.append(f"Error in the find_product_notice method: {e}")
            return np.nan

    @staticmethod
    def find_quantity(text):
        """
        Method for determining the product's quantity
        """
        try:
            quantity_match = re.search(r'(?i)кол(\s*-\s*|ичест)?во\s*(?:\D*)(:)?\s*(\d+)', text)
            quantity = int(quantity_match.group(3)) if quantity_match else np.nan
            return quantity

        except Exception as e:
            self.error_log.append(f"Error in the find_quantity method: {e}")
            return np.nan

    def find_unit_note(self, text):
        """
        Method for determining the quantity's unit
        """
        try:
            unit_match = re.search(r'(?i)кол(-|ичест)?во\s*(?:\D*)(:)?\s*(\d+)\s*(\D+)\s*\\n\d', text)
            unit = unit_match.group(4).strip() if unit_match else np.nan

            if unit:
                unit_index = self.units_df[self.units_df['data'] == unit].index
                if not unit_index.empty:
                    unit_found = self.units_df.loc[unit_index[0], 'unit']
                    return unit_found
            return np.nan

        except Exception as e:
            self.error_log = pd.concat(
                [self.error_log, pd.DataFrame({'Ошибка': [f"Error in the find_unit_note method: {e}"],
                                               'Заявка': [text]})], ignore_index=True)
            return np.nan

    @staticmethod
    def find_car(text):
        """
        Method for determining the car's numbers
        """
        try:
            car_match = re.findall(r'(?i)[А-Я]{1}\d{3}[А-Я]{2}\d{3}\s*\w*', text)
            cars = car_match if car_match else np.NAN
            return cars

        except Exception as e:
            self.error_log.append(f"Error in the find_purchaser method: {e}")
            return np.nan

    @staticmethod
    def find_organization(text):
        """
        Method for determining our organization's name
        """
        try:
            organization_match = re.search(r'(?i)(продажа\s+от:|(продажа)?\s*от\s+(клиента)?\s*(:)?)\s*(.+?)\\n', text)
            organization = organization_match.group(5) if organization_match else np.NAN
            return organization

        except Exception as e:
            self.error_log.append(f"Error in the find_organization method: {e}")
            return np.nan

    @staticmethod
    def find_transshipment(text):
        """
        Method for determining the transshipment's point
        """
        try:
            transshipment_match = \
                re.search(r'(?i)\d+\.\s*(С\s+(перевалки)?\s*|Завод\s*(отгрузки)?\s*(:)?|Перевалка\s*(:)?)\s*(.+?)\\n',
                          text)
            transshipment = transshipment_match.group(6) if transshipment_match else np.nan
            return transshipment

        except Exception as e:
            self.error_log.append(f"Error in the find_transshipment method: {e}")
            return np.nan

    @staticmethod
    def find_purchaser(text):
        """
        Method for determining the purchaser's name
        """
        try:
            purchaser_match = re.search(r'(?i)\d+\.\s*(Покупатель\s*(груза)?\s*(:)?)\s*(.+?)\\n', text)
            purchaser = purchaser_match.group(4) if purchaser_match else np.nan
            return purchaser

        except Exception as e:
            self.error_log.append(f"Error in the find_purchaser method: {e}")
            return np.nan

    @staticmethod
    def find_consignee(text):
        """
        Method for determining the consignee's name
        """
        try:
            consignee_match = re.search(
            r'(?i)\d+\.\s*(?:Грузопол\w*\s*(?::)?|Грузопол\w*\s*\(при\s*оформ\w*\s*ттн\)\s*(?::)?)\s*(.+?)\\n',
            text)
            consignee = consignee_match.group(1).split(':')[-1].strip() if consignee_match else np.NAN
            return consignee

        except Exception as e:
            self.error_log.append(f"Error in the find_purchaser method: {e}")
            return np.nan

    @staticmethod
    def find_consignee_leg_addr(text):
        """
        Method for determining the legal consignee's address
        """
        try:
            find_consignee_leg_addr_match = re.search(
            r'(?i)\d+\.\s*(юр\w*\s*(?:\.)?\s*адрес\s*грузополучателя|адрес\s*грузополучателя\s*\(юр\w*\s*(?:\.)?\))\s*(?::)?\s*(.+?)\\n',
            text)
            find_consignee_leg_addr = find_consignee_leg_addr_match.group(2) if find_consignee_leg_addr_match else np.NAN
            return find_consignee_leg_addr

        except Exception as e:
            self.error_log.append(f"Error in the find_consignee_leg_addr method: {e}")
            return np.nan

    def find_unl_addr(self, text):
        """
        Method for determining unload addresses
        """
        try:
            for index, row in self.unload_addresses_df.iterrows():
                pattern1 = row['name']
                pattern2 = row['place']
                address = row['address']

                if pattern1.lower().replace(' ', '') and pattern2.lower().replace(' ', '') \
                        in text.lower().replace(' ', ''):
                    return address
            return np.nan

        except Exception as e:
            self.error_log.append(f"Error in the find_unl_addr method: {e}")
            return np.nan

    @staticmethod
    def find_phones(text):
        """
        Method for determining the phone numbers
        """
        try:
            phone_numbers = re.findall(r'\+?\d{1,3}[\s-]?\(?\d{3}\)?[\s-]?\d{2,3}[\s-]?\d{2}[\s-]?\d{2}', text)
            phone_numbers = phone_numbers if phone_numbers else np.NAN
            return phone_numbers

        except Exception as e:
            self.error_log.append(f"Error in the find_phones method: {e}")
            return np.nan

    @staticmethod
    def find_time(text):
        """
        Method for determining the unloading time
        """
        try:
            time_match = re.search(r'(?i)(время)?\s*при(ё|е)мк(и|а)(?::)?\s*(.*?)\s*\\n', text)
            time = time_match.group(4) if time_match else np.NAN
            return time

        except Exception as e:
            self.error_log.append(f"Error in the find_time method: {e}")
            return np.nan

    @staticmethod
    def find_note(text):
        """
        Method for determining the note
        """
        try:
            if 'оплата' in text.lower():
                note_match = re.search(r'(?i)(оплата)\s*(?::)?\s*(.*?)\\n', text)
                note = '{} {}'.format(note_match.group(1).strip(), note_match.group(2).strip() if note_match else np.nan)
                return note
            else:
                return np.nan

        except Exception as e:
            self.error_log.append(f"Error in the find_note method: {e}")
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
            delivery_found = self.delivery_type(application)
            product_found = self.find_product(application)
            product_notice_found = self.find_product_notice(application)
            quantity_found = self.find_quantity(application)
            unit_found = self.find_unit_note(application)
            num_str = round(quantity_found / 35 if unit_found == 'т' else quantity_found / 40)
            cars_found = ', '.join(self.find_car(application)) if self.find_car(application) is not np.nan else np.nan
            organization_found = self.find_organization(application)
            transshipment_found = self.find_transshipment(application)
            purchaser_found = self.find_purchaser(application)
            consignee_found = self.find_consignee(application)
            consignee_leg_addr_found = self.find_consignee_leg_addr(application)
            unl_addr_found = self.find_unl_addr(application)
            phones_found = ', '.join(self.find_phones(application)) if self.find_phones(application) is not np.nan else np.nan
            accept_time_found = self.find_time(application)
            note_found = self.find_note(application)


            new_row = pd.DataFrame({'Менеджер': [manager_found],
                                    'Дата': [formatted_date],
                                    'Вид доставки': [delivery_found],
                                    'Товар': [product_found],
                                    'Примечание к Товару': [product_notice_found],
                                    'Кол-во': [35 if unit_found == 'т' else 40],
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
