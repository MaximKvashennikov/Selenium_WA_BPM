import pandas as pd
from pandas import concat
import pretty_html_table
import os
import time
import datetime
import pandas.io.formats.excel
from pathlib import Path


class GetDataTable:
    def __init__(self, mr_name=None, path_file_sw=None):
        self.mr_name = mr_name
        self.path_file_sw = path_file_sw

    def read_ex(self):
        # path_file = "\\".join(os.getcwd().split("\\")[:-1])

        """Основной путь"""
        path_file = os.getcwd()
        # print(path_file)
        pd.set_option('display.max_colwidth', 2000)
        df = pd.read_excel(
            path_file + "\\Входные данные.xlsx",
            encoding='cp1251',
        )

        return df

    def get_responsible(self):
        return self.read_ex().iloc[0][0]

    def get_executor(self):
        return self.read_ex().iloc[0][1]

    def get_start_date(self):
        return self.read_ex().iloc[0][2].split('"')[1]

    def get_end_date(self):
        return self.read_ex().iloc[0][3].split('"')[1]

    def get_reg_list(self):
        return self.read_ex()['МР - Регион (как в отчете SW расширения WFL)'].tolist()

    def read_sw(self):
        path_file = self.path_file_sw
        pd.set_option('display.max_colwidth', 2000)
        df = pd.read_excel(
            path_file,
            encoding='cp1251',
        )

        return df

    def get_column_sw_file(self):
        df2 = self.read_sw()
        df2 = df2[['МР - Р', 'Имя элемента', 'Влияние']]

        # Берем РРЛ и влияние
        df3 = df2[df2['МР - Р'] == self.mr_name]
        df_rrl = df3['Имя элемента'].tolist()
        df_influence = df3['Влияние'].tolist()

        return df_rrl, df_influence


def main():
    print(GetDataTable().read_ex())
    print(GetDataTable().get_responsible())
    print(GetDataTable().get_executor())
    print(GetDataTable().get_start_date())
    print(GetDataTable().get_end_date())


if __name__ == "__main__":
    # main()
    pass
