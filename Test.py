#!/usr/bin/env python

import requests 
import pandas as pd
import sys
import re
import os
import win32com.client as win32

def get_company_list() :
#copy csv data to xlsx
	lastCol = mWorkSheet.UsedRange.Columns.Count
	lastRow = mWorkSheet.UsedRange.Rows.Count
	print("Laster Row : " + str(lastRow) + " colum num : " + str(lastCol))
	#mWorkSheet.Range("B2:B" + str(lastRow)).Select()  # 기업코드

	for i in range (2, 12) :  #lastRow) :
		url = 'https://finance.naver.com/item/sise.nhn?code=' + str(mWorkSheet.Cells(i,2).Value)
		print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함
		table_df = table_df_list[1]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
		df = pd.DataFrame(table_df)
		# df.loc = row, column, 0부터 시작
		mWorkSheet.Cells(i, 4).Value = df.loc[0][1] #현재가
		#print(df.head()) #print dataframe data
		#print(df.shape) #get row, column count

		url = 'https://finance.naver.com/item/main.nhn?code=' + str(mWorkSheet.Cells(i,2).Value)
		print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')
		table_df = table_df_list[5] #주요 재무제표
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 5).Value = df.loc[2][1] #상장주식수
		mWorkSheet.Cells(i, 6).Value = "=D"+ str(i) + "*E" + str(i) + "/100000000"   #시가총액(억)

		table_df = table_df_list[3] #주요 재무제표
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 22).Value = df.loc[0][1] # 2017.12 (Y) 매출액
		mWorkSheet.Cells(i, 23).Value = df.loc[0][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 24).Value = df.loc[0][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 25).Value = df.loc[0][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 27).Value = df.loc[0][5] # 2019.06
		mWorkSheet.Cells(i, 28).Value = df.loc[0][6] # 2019.09
		mWorkSheet.Cells(i, 29).Value = df.loc[0][7] # 2019.12
		mWorkSheet.Cells(i, 30).Value = df.loc[0][8] # 2020.03
		mWorkSheet.Cells(i, 31).Value = df.loc[0][9] # 2020.06
		mWorkSheet.Cells(i, 32).Value = df.loc[0][10] # 2020.09 (E)

		mWorkSheet.Cells(i, 34).Value = df.loc[1][1] # 2017.12 (Y) 영업이익
		mWorkSheet.Cells(i, 35).Value = df.loc[1][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 36).Value = df.loc[1][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 37).Value = df.loc[1][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 39).Value = df.loc[1][5] # 2019.06
		mWorkSheet.Cells(i, 40).Value = df.loc[1][6] # 2019.09
		mWorkSheet.Cells(i, 41).Value = df.loc[1][7] # 2019.12
		mWorkSheet.Cells(i, 42).Value = df.loc[1][8] # 2020.03
		mWorkSheet.Cells(i, 43).Value = df.loc[1][9] # 2020.06
		mWorkSheet.Cells(i, 44).Value = df.loc[1][10] # 2020.09 (E)

		mWorkSheet.Cells(i, 46).Value = df.loc[2][1] # 2017.12 (Y) 당기순이익
		mWorkSheet.Cells(i, 47).Value = df.loc[2][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 48).Value = df.loc[2][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 49).Value = df.loc[2][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 51).Value = df.loc[2][5] # 2019.06
		mWorkSheet.Cells(i, 52).Value = df.loc[2][6] # 2019.09
		mWorkSheet.Cells(i, 53).Value = df.loc[2][7] # 2019.12
		mWorkSheet.Cells(i, 54).Value = df.loc[2][8] # 2020.03
		mWorkSheet.Cells(i, 55).Value = df.loc[2][9] # 2020.06
		mWorkSheet.Cells(i, 56).Value = df.loc[2][10] # 2020.09 (E)

		table_df = table_df_list[4]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 7).Value = df.loc[4][1] #외국인비율(%)

		mWorkSheet.Cells(i, 58).Value = "=AC" + str(i) + "/Q" + str(i)   #영업이익률(2017.12)
		mWorkSheet.Cells(i, 59).Value = "=AD" + str(i) + "/R" + str(i)   #영업이익률(2018.12)
		mWorkSheet.Cells(i, 60).Value = "=AE" + str(i) + "/S" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 61).Value = "=AF" + str(i) + "/T" + str(i)   #영업이익률(2020.12)

		mWorkSheet.Cells(i, 63).Value = "=AH" +str(i) + "/V" + str(i)   #영업이익률(2019.06)
		mWorkSheet.Cells(i, 64).Value = "=AI" +str(i) + "/W" + str(i)   #영업이익률(2019.09)
		mWorkSheet.Cells(i, 65).Value = "=AJ" +str(i) + "/X" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 66).Value = "=AK" +str(i) + "/Y" + str(i)   #영업이익률(2020.03)
		mWorkSheet.Cells(i, 67).Value = "=AL" +str(i) + "/Z" + str(i)   #영업이익률(2020.06)
		mWorkSheet.Cells(i, 68).Value = "=AM" +str(i) + "/AA" + str(i)  #영업이익률(2020.09)

		mWorkSheet.Cells(i, 70).Value = "=AO" +str(i) + "/Q" + str(i)   #순이익률(2017.12)
		mWorkSheet.Cells(i, 71).Value = "=AP" +str(i) + "/R" + str(i)   #순이익률(2018.12)
		mWorkSheet.Cells(i, 72).Value = "=AQ" +str(i) + "/S" + str(i)   #순이익률(2019.12)
		mWorkSheet.Cells(i, 73).Value = "=AR" +str(i) + "/T" + str(i)   #순이익률(2020.12)

		mWorkSheet.Cells(i, 75).Value = "=AT" +str(i) + "/V" + str(i)   #영업이익률(2019.06)
		mWorkSheet.Cells(i, 76).Value = "=AU" +str(i) + "/W" + str(i)   #영업이익률(2019.09)
		mWorkSheet.Cells(i, 77).Value = "=AV" +str(i) + "/X" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 78).Value = "=AW" +str(i) + "/Y" + str(i)   #영업이익률(2020.03)
		mWorkSheet.Cells(i, 79).Value = "=AX" +str(i) + "/Z" + str(i)   #영업이익률(2020.06)
		mWorkSheet.Cells(i, 80).Value = "=AY" +str(i) + "/AA" + str(i)  #영업이익률(2020.09)

		mWorkSheet.Cells(i, 80).Value = "=AY" +str(i) + "/AA" + str(i)  #PER = 주가 / 주당순이익
		


#def run_each_company_data(company_code) :

#read file
filename = "export_Data.xlsx"
filepath = os.path.abspath(filename)
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

mWorkBook = excel.Workbooks.Open(filepath)
mWorkSheet = mWorkBook.Worksheets('RawData')
mWorkSheet.Select()
get_company_list()
#url = 'https://finance.naver.com/item/main.nhn?code=005930'
#table_df_list = pd.read_html(url, encoding='euc-kr')
#table_df = table_df_list[5]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
#df = pd.DataFrame(table_df)
#print('table_df_list[5]')
#print(df.head()) #print dataframe data


mWorkBook.Save()
mWorkBook.Close()