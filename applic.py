import pandas as pd
import numpy as np
import re
import traceback


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
        except Exception as e:
            self.error_log.append(f"Error in the find_manager method: {e}")
            return np.nan

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
            self.error_log.append(f"Ошибка в методе delivery_type: {e}")
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
            self.error_log.append(f"Ошибка в методе find_quantity: {e}")
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
                [self.error_log, pd.DataFrame({'Ошибка': [f"Ошибка в методе find_unit_note: {e}"],
                                               'Заявка': [text]})],
                ignore_index=True)
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
            self.error_log.append(f"Ошибка в методе find_purchaser: {e}")
            return np.nan

    def process_application(self, application):
        """
        Method to process an application and update the dataframe
        """
        try:
            manager_found = self.find_manager(application)
            delivery_found = self.delivery_type(application)
            product_found = self.find_product(application)
            product_notice_found = self.find_product_notice(application)
            quantity_found = self.find_quantity(application)
            unit_found = self.find_unit_note(application)
            purchaser_found = self.find_purchaser(application)

            new_row = pd.DataFrame({'Менеджер': [manager_found],
                                    'Вид доставки': [delivery_found],
                                    'Товар': [product_found],
                                    'Примечание к Товару': [product_notice_found],
                                    'Кол-во': [quantity_found],
                                    'Ед.изм.': [unit_found],
                                    'Покупатель': [purchaser_found],
                                    'Текст заявки': [application]})

            self.applications_df = pd.concat([self.applications_df, new_row], ignore_index=True)

        except Exception as e:
            traceback_str = traceback.format_exc()
            self.error_log = pd.concat([self.error_log, pd.DataFrame({'Ошибка': [traceback_str],
                                                                      'Заявка': [application]})],
                                       ignore_index=True)

            new_row = pd.DataFrame({'Текст заявки': [application]})
            self.applications_df = pd.concat([self.applications_df, new_row], ignore_index=True)

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
