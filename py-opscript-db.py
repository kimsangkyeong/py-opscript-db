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

pd.set_option("display.max_colwidth", 100)  # 컬럼 정보 보여주기
pd.set_option("display.max_rows", 999)  # row 정보 보여주기

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
                                       'font_size': 9,
                                       'text_wrap': True, 
                                       'fg_color': '#D9D9D9', 
                                       'border': 1})
  col_width_list = []
  for col_num, value in enumerate(df.columns.values):
    worksheet.write(2, col_num, value, header_format)
    if value in ["script_user", "sspord_grant", "sspstl_grant", "sspcmp_grant"] :
      col_width_list.append(35)
    else :
      col_width_list.append(7) # initialize

  # Set the Column Width
  for col, colwidth in enumerate(col_width_list):
    worksheet.set_column(col,col+1, colwidth)
  # Add the border to the table
  data_format = workbook.add_format({'align': 'left', 
                                     'valign': 'vcenter', 
                                     'text_wrap': True,
                                     'font_size': 8,
                                     'border': 1})
  missing_data_format = workbook.add_format({'align': 'left', 
                                     'valign': 'vcenter', 
                                     'text_wrap': True,
                                     'font_size': 8,
                                     'fg_color': '#F79646', 
                                     'border': 1})
  row_idx, col_idx = df.shape
  for row in range(row_idx):
    for col in range(col_idx):
      if df.values[row, col] == "MISSING" :
        worksheet.write(row + 3, col, df.values[row, col], missing_data_format)
      else :
        worksheet.write(row + 3, col, df.values[row, col], data_format)

  # Add the remark to the excel
  worksheet.write(len(df)+4, 0, 'Remark:', workbook.add_format({'bold': True}))
  worksheet.write(len(df)+5, 0, 'The last update time is ' + datetime.now().strftime('%Y-%m-%d %H:%M') + '.')

def data_preprocessing(indf):

  df = indf

  # P사번 컬럼 숫자형을 string type을 강제 설정
  df["P사번"] = df["P사번"].astype('string')

  # 결측치 None으로 대체
  df["P사번"] = df["P사번"].fillna("MISSING")
  df["대상DB"] = df["대상DB"].fillna("MISSING")
  df["대상 스키마"] = df["대상 스키마"].fillna("MISSING")
  df["권한"] = df["권한"].fillna("MISSING")

  # 대문자로 표준화
  df.loc[:,"P사번"] = df["P사번"].replace(' ','',regex=True).str.upper()
  df.loc[:,"대상DB"] = df["대상DB"].replace(' ','',regex=True).str.upper()
  df.loc[:,"대상 스키마"] = df["대상 스키마"].replace(' ','',regex=True).str.upper()
  df.loc[:,"권한"] = df["권한"].replace(' ','',regex=True).str.upper()
  
  # script 결과 저장 추가 컬럼 추가
  df["script_user"] = ""
  df["sspord_grant"] = ""
  df["sspstl_grant"] = ""
  df["sspcmp_grant"] = ""
  return df

def grant_script(listschema, userid, dml_str, dbinfo, schemadf):
  grant_list = []
  print(f'grant_script : {userid}, {dbinfo}, {listschema}, {dml_str}')
  for schemaitem in listschema :
    if schemaitem != "MISSING" :
      if ( "전체" in schemaitem or "ALL" in schemaitem ) :
        for dbidx in range(schemadf.shape[0]):
          if dbinfo == (schemadf.loc[dbidx, "DB"]).upper() :
            for schemainfo in schemadf.loc[dbidx, "SCHEMA"].split("/") :
              grant_list.append(f"grant {dml_str} on {schemainfo}.* to '{userid}'@'%'; ")
      else :
        grant_list.append(f"grant {dml_str} on {schemaitem}.* to '{userid}'@'%'; ")
  grant_str = ''.join(grant_list)
  # print(f'grant_str : {grant_str}')
  return grant_str

def generate_script(datadf, schemadf, envdf):
  gendf = datadf 
  initpasswd = envdf.loc[0,'InitPasswd'] 
  dml_keymap = {"C" : "insert", "R" : "select", "U" : "update", "D" : "delete",
                "A" : "insert,select,update,delete" }
  dbinfo_keymap = {"SSPORD" : "sspord_grant", "SSPSTL" : "sspstl_grant", "SSPCMP" : "sspcmp_grant"}

  for rowidx in range(gendf.shape[0]):
    # print(f'idx = {rowidx}, {df.loc[rowidx, cols]}')
    usernm = gendf.loc[rowidx, "사용자"]
    userid = gendf.loc[rowidx, "P사번"]
    dbinfo = gendf.loc[rowidx, "대상DB"]
    rawschema = gendf.loc[rowidx, "대상 스키마"]
    listschema = rawschema.split("/")
    rawauthority = gendf.loc[rowidx, "권한"]
    listauthority = rawauthority.split("/")
    # print(f'idx = {rowidx},{usernm},{userid},{dbinfo},{rawschema},{listschema},{rawauthority},{listauthority}')

    # create user script 생성
    if userid != "MISSING" :
      script_user = f"create user '{userid}'@'%' identified by '{initpasswd}';" 
      gendf.loc[rowidx,"script_user"] = script_user
    
    # Grant 권한 문자열 생성
    if rawauthority != "MISSING" :
      if ( "전체" in rawauthority or "ALL" in rawauthority ):
        dml_str = dml_keymap["A"]
      else :
        dml_str_list = []
        for auth in listauthority:
            if auth != "MISSING" :
              dml_str_list.append(dml_keymap[auth])
        dml_str = ','.join(dml_str_list)
    else :
      dml_str = "MISSING" 
    # print(f'dml_str : {dml_str}')

    # SSPORD/SSPSTL/SSPCMP DBMS 권한부여 
    if not ( rawschema == "MISSING" or dml_str == "MISSING" or dbinfo == "MISSING" ) :
      gendf.loc[rowidx,dbinfo_keymap[dbinfo]] = grant_script(listschema, userid, dml_str, dbinfo, schemadf)

  return gendf

def main(argv):

  # 파일 읽어오기
  raw_df = pd.read_excel(argv[1], sheet_name='request')
  # print(f'{raw_df}')
  schemadf = pd.read_excel(argv[1], sheet_name='schema')
  # print(f'{schemadf.head()}')
  envdf = pd.read_excel(argv[1], sheet_name='env')
  # print(f'{envdf.head()}')

  # 데이타 전처리 : Cleansing
  datadf = data_preprocessing(raw_df)
  # print(f'==> {datadf}')
  
  # script sql 문장 생성하기
  gendf = generate_script(datadf, schemadf, envdf)

  # 현재 디렉토리
  path_cwd = os.getcwd()
  # OS 판단  : win32, linux, cygwin, darwin, aix
  my_os = sys.platform
  if my_os == "linux":
    output_file = f'{path_cwd}/generated_dbscript_{datetime.now().strftime("%Y%m%d-%H%M")}.xlsx'
  else:
    output_file = f'{path_cwd}\generated_dbscript_{datetime.now().strftime("%Y%m%d-%H%M")}.xlsx'

  # print(f'{gendf}')

  with pd.ExcelWriter(output_file, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
    df_to_excel(writer, gendf, 'dbscripts')  

  print("작업완료 : 생성된 output_file : ", output_file)

if __name__ == "__main__":
   main(sys.argv[:])
