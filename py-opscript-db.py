####################################################################################################
# 
# Purpose   : main program for db operation script generation
# Source    : py-opscript-db.py
# Usage     : python py-opscript-db.py 
# Developer : ksk
# --------  -----------   -------------------------------------------------
# Version :     date    :  reason
#  1.0      2023.11.22     first create
#
# ref     : https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html, .to_excel.html
# ref     : https://towardsdatascience.com/use-python-to-stylize-the-excel-formatting-916e00e33302
#           https://xlsxwriter.readthedocs.io/index.html
# required install package : # pip install numpy, pandas, xlsxwriter, openpyxl, ipython
#
####################################################################################################
### This first line is for modules to work with Python 2 or 3
from __future__ import print_function
from distutils.log import debug
import os
import sys, getopt
import argparse
import importlib
import numpy as np   # pip install numpy
import pandas as pd  # pip install pandas
from datetime import date, datetime
import IPython.display as display # pip install ipython

pd.set_option("display.max_colwidth", 999)  # 컬럼 정보 보여주기
pd.set_option("display.max_rows", 150)  # row 정보 보여주기

def df_to_excel(writer, df, sheetname):
  df.to_excel(writer, sheet_name=sheetname, startrow = 2, index=False)
  # Add the title to the excel
  workbook = writer.book
  worksheet = writer.sheets[sheetname]
  worksheet.write(0, 0, f'[ {sheetname} ]',
                  workbook.add_format({'bold': True, 'color': '#E26B0A', 'size': 14}))
  # Add the color to the table header
  header_format = workbook.add_format({'bold': True, 
                                       'align': 'center', 
                                       'valign': 'vcenter', 
                                       'font_size': 10,
                                       'text_wrap': True, 
                                       'fg_color': '#D9D9D9', 
                                       'border': 1})
  col_width_list = []
  for col_num, value in enumerate(df.columns.values):
    worksheet.write(2, col_num, value, header_format)
    col_width_list.append(12) # initialize

  # klogger_dat.debug(f'col_width_list : {col_width_list}')
  # Set the Column Width
  for col, colwidth in enumerate(col_width_list):
    worksheet.set_column(col,col+1, colwidth)
  # Add the border to the table
  data_format = workbook.add_format({'align': 'left', 
                                     'valign': 'vcenter', 
                                     'text_wrap': True,
                                     'font_size': 9,
                                     'border': 1})
  row_idx, col_idx = df.shape
  for row in range(row_idx):
    for col in range(col_idx):
      if type(df.values[row, col]) == type(dict()) :
        data_len = np.ceil(len(str(df.values[row, col])) / 4)
        if col_width_list[col] < data_len:
          if data_len >= 100 :
            col_width_list[col] = 100
          else :
            col_width_list[col] = data_len
          worksheet.set_column(col,col+1, col_width_list[col])
        worksheet.write(row + 3, col, str(df.values[row, col]), data_format)
      elif type(df.values[row, col]) == type(str()) : 
        data_len = len(df.values[row, col]) 
        if col_width_list[col] < data_len :
          if data_len >= 60 :
            col_width_list[col] = 50
          elif data_len >= 40 :
                col_width_list[col] = 40
          elif data_len >= 30 :
            col_width_list[col] = 25
          else :
            col_width_list[col] = data_len
          worksheet.set_column(col,col+1, col_width_list[col])
        worksheet.write(row + 3, col, df.values[row, col], data_format)
      else :
        worksheet.write(row + 3, col, df.values[row, col], data_format)

  # Merge Cell in column 'A' only
  # Merge cell format
  merge_format = workbook.add_format({'font_size': 9,
                                      'border': 1,
                                      'align': 'left',
                                      'valign': 'vcenter',
                                      'text_wrap': True})
  col = 0; # column 'A' 
  befdata = ''; startrow = 0; currrow = 0;
  for row in range(row_idx):
    currrow = row
    if befdata == df.values[row, col] :
      continue
    else :
      if befdata != '' and (row - startrow >= 2) :
        # klogger.debug(f'startrow : [{startrow}], row : [{row}], befdata : [{befdata}], data : [{df.values[row, col]}]')
        worksheet.merge_range(f'A{startrow+4}:A{(row-1)+4}', str(befdata) if type(befdata) == type(dict()) else befdata, merge_format)
      startrow = row
      befdata = df.values[row, col]
  # klogger.debug(f'for loop : startrow : [{startrow}], row : [{currrow}], befdata : [{befdata}], data : [{df.values[row, col]}]')
  if currrow - startrow >= 1 : # last merge cell check
    worksheet.merge_range(f'A{startrow+4}:A{currrow+4}', str(befdata) if type(befdata) == type(dict()) else befdata, merge_format)
    
  # Add the remark to the excel
  worksheet.write(len(df)+4, 0, 'Remark:', workbook.add_format({'bold': True}))
  worksheet.write(len(df)+5, 0, 'The last update time is ' + datetime.now().strftime('%Y-%m-%d %H:%M') + '.')

def main(argv):
  print(f'argv : {argv}')
  df = pd.read_excel(argv[1], sheet_name='request')
  print(f'{df.columns}')
  print(f'{df.info()}')
  # 결측치 None으로 대체
  df["P사번"] = df["P사번"].fillna("MISSING")
  df["대상DB"] = df["대상DB"].fillna("MISSING")
  df["대상 스키마"] = df["대상 스키마"].fillna("MISSING")
  df["권한"] = df["권한"].fillna("MISSING")

  # 대문자로 표준화
  df.loc[:,"대상DB"] = df["대상DB"].str.upper()
  df.loc[:,"대상 스키마"] = df["대상 스키마"].str.upper()
  df.loc[:,"권한"] = df["권한"].str.upper()
  cols = ["사용자","P사번", "대상DB", "대상 스키마", "권한"]

  # 추가 컬럼 추가
  df["script_user"] = ""
  df["sspord_grant"] = ""
  df["sspstl_grant"] = ""
  df["sspcmp_grant"] = ""
  excols = ["사용자","P사번", "대상DB", "대상 스키마", "권한","script_user","sspord_grant","sspstl_grant","sspcmp_grant"]
#   print(f'{df[cols]}')
  print(f'{df.shape}, rowcount = {df.shape[0]}')
  initpasswd="sktssp20~!"
  for rowidx in range(df.shape[0]):
    # print(f'idx = {rowidx}, {df.loc[rowidx, cols]}')
    usernm = df.loc[rowidx, "사용자"]
    userid = df.loc[rowidx, "P사번"]
    dbinfo = df.loc[rowidx, "대상DB"]
    rawschema = df.loc[rowidx, "대상 스키마"]
    listschema = rawschema.split("/")
    rawauthority = df.loc[rowidx, "권한"]
    listauthority = rawauthority.split("/")
    print(f'idx = {rowidx},{usernm},{userid},{dbinfo},{rawschema},{listschema},{rawauthority},{listauthority}')
    # create user script 생성
    if userid != "MISSING" :
      print(f'userid {userid}, type{type(userid)}')
      script_user = f"create user '{userid}'@'%' identified by '{initpasswd}';" 
      df.loc[rowidx,"script_user"] = script_user
    # SSPORD DBMS 권한부여
    if dbinfo == "SSPORD" :
      df.loc[rowidx,"sspord_grant"] = "SSPORD"
    # SSPSTL DBMS 권한부여
    if dbinfo == "SSPSTL" :
      df.loc[rowidx,"sspstl_grant"] = "SSPSTL"
    # SSPCMP DBMS 권한부여
    if dbinfo == "SSPCMP" :
      df.loc[rowidx,"sspcmp_grant"] = "SSPCMP"

    if rowidx > 8 :
      break
#   with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
#     df_to_excel(writer, df_route53           , 'route53')                     # 1
  print(f'{df[excols].head()}')

if __name__ == "__main__":
   main(sys.argv[:])
