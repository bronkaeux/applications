import pandas as pd
import numpy as np
import re


class ApplicationProcessor:
    def __init__(self, managers_df, units_df):
        self.managers_df = managers_df
        self.units_df = units_df
        self.df = pd.DataFrame(columns=['Менеджер', 'Вид доставки', 'Кол-во', 'Ед.изм.', 'Покупатель', 'Текст заявки'])
        self.error_log = []

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
            self.error_log.append(f"Ошибка в методе find_manager: {e}")
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
            unit_found = self.units_df.loc[self.units_df[self.units_df['data'] == unit].index[0], 'unit']
            return unit_found
        except Exception as e:
            self.error_log.append(f"Ошибка в методе find_unit_note: {e}")
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
            quantity_found = self.find_quantity(application)
            unit_found = self.find_unit_note(application)
            purchaser_found = self.find_purchaser(application)

            new_row = pd.DataFrame({'Менеджер': [manager_found],
                                    'Вид доставки': [delivery_found],
                                    'Кол-во': [quantity_found],
                                    'Ед.изм.': [unit_found],
                                    'Покупатель': [purchaser_found],
                                    'Текст заявки': [application]})

            self.df = pd.concat([self.df, new_row], ignore_index=True)

        except Exception as e:
            self.error_log.append(f"Ошибка при обработке заявки: {e}")
            new_row = pd.DataFrame({'Менеджер': [application]})
            self.df = pd.concat([self.df, new_row], ignore_index=True)

    def save_to_excel(self, filename):
        """
        Method to save the dataframe to Excel
        """
        with pd.ExcelWriter(filename) as writer:
            self.df.to_excel(writer, sheet_name="data", index=False)

        if self.error_log:
            with open("error_log.txt", "w") as file:
                file.write("\n".join(self.error_log))


# loading dictionaries
managers = pd.DataFrame({'login': ['ИГО', 'Юра Менеджер', 'Алексей Мельхер'],
                         'manager': ['Игорь Хабаров', 'Юрий', 'Алексей Мельхер']})

units = pd.DataFrame({'data': ['т', 'т.', 'тн', 'тонн', 'кубов'],
                      'unit': ['т', 'т', 'т', 'т', 'куб']})

# entering an application
application = str(input('Введите заявку: '))

# Create an instance of ApplicationProcessor
processor = ApplicationProcessor(managers, units)

# Process the application
processor.process_application(application)

# Save to Excel
processor.save_to_excel("reg.xlsx")
