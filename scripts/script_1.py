# pip install XlsxWriter pandas xlsxwriter xlrd xlwt
import requests
import pandas as pd
import xlrd
import xlwt
import xlsxwriter
import openpyxl
from openpyxl import Workbook

# response = requests.get('https://api.marketinout.com/run/screen?key=334ebaa78b234e7c')
# text = response.text
# content = response.content

# with open("response.txt", "w", encoding='utf8') as f:
#     f.write(response.text)  # Записываем ответ от сервера в текстовый файл
#     f.flush()  # Сохраняем файл на жесткий диск
#     f.close()  # Закрываем файл
#     print('ОК...response.txt перезаписан')

# txt_report = open('response.txt', 'r', encoding='utf8')  # Откроем файл для чтения
#
# for line in txt_report:
#     list_ = txt_report.readline().removesuffix('\n').split('|')
#     print(list_)


df = pd.DataFrame({
    'symbol'                   : [],
    'last'                     : [],
    'date'                     : [],
    'altman_z_score'           : [],
    'debt_equity_ratio'        : [],
    'interest_cover'           : [],
    'current_ratio'            : [],
    'piotroski_f_score'        : [],
    'beneish_m_score'          : [],
    'price_intrinsic_potential': [],
    'price_lynch_potential'    : [],
    'price_graham_potential'   : [],
    'shiller_pe_ratio'         : [],
    'roa'                      : [],
    'roa_5y_avg'               : [],
    'roe'                      : [],
    'roe_5y_avg'               : [],
    'name'                     : []
})

count = 0
with open('response.txt') as file:
    for line in file:
        list_ = line.removesuffix('\n').split('|')
        if len(list_) > 1:
               df_new_row = pd.DataFrame({
                   'symbol'                   : [list_[0].removesuffix('.ME')],
                   'last'                     : [list_[1]],
                   'date'                     : [list_[2]],
                   'altman_z_score'           : [list_[3]],
                   'debt_equity_ratio'        : [list_[4]],
                   'interest_cover'           : [list_[5]],
                   'current_ratio'            : [list_[6]],
                   'piotroski_f_score'        : [list_[7]],
                   'beneish_m_score'          : [list_[8]],
                   'price_intrinsic_potential': [list_[9]],
                   'price_lynch_potential'    : [list_[10]],
                   'price_graham_potential'   : [list_[11]],
                   'shiller_pe_ratio'         : [list_[12]],
                   'roa'                      : [list_[13]],
                   'roa_5y_avg'               : [list_[14]],
                   'roe'                      : [list_[15]],
                   'roe_5y_avg'               : [list_[16]],
                   'name'                     : [0]})
               df = pd.concat([df, df_new_row])
               count += 1
print(f'Записано строк: {count}')

int_list = [
    'last'                      ,
    'altman_z_score'            ,
    'debt_equity_ratio'         ,
    'interest_cover'            ,
    'current_ratio'             ,
    'piotroski_f_score'         ,
    'beneish_m_score'           ,
    'price_intrinsic_potential' ,
    'price_lynch_potential'     ,
    'price_graham_potential'    ,
    'shiller_pe_ratio'          ,
    'roa'                       ,
    'roa_5y_avg'                ,
    'roe'                       ,
    'roe_5y_avg'                ,
]
df[int_list] = df[int_list].apply(pd.to_numeric)

# Указать writer библиотеки
writer = pd.ExcelWriter('2023.XX.XX_russian_stocks.xlsx', engine='xlsxwriter')
df.to_excel(writer, 'Sheet1')  # Записать ваш DataFrame в файл
writer.close()  # Сохраним результат


