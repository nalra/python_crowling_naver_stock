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

	for i in range (2, 11) : #lastRow) :
		url = 'https://finance.naver.com/item/sise.nhn?code=' + str(mWorkSheet.Cells(i,2).Value)
		print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함
		table_df = table_df_list[1]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
		df = pd.DataFrame(table_df)
		# df.loc = row, column, 0부터 시작
		mWorkSheet.Cells(i, 8).Value = df.loc[0][1] #현재가
		mWorkSheet.Cells(i, 9).Value = str(df.loc[1][1]) + " " + str(df.loc[2][1]) #전일대비 + 등락률(%)
		#print(df.head()) #print dataframe data
		#print(df.shape) #get row, column count

		url = 'https://finance.naver.com/item/main.nhn?code=' + str(mWorkSheet.Cells(i,2).Value)
		#print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')
		table_df = table_df_list[5] #주요 재무제표
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 10).Value = df.loc[2][1] #상장주식수
		mWorkSheet.Cells(i, 11).Value = "=H"+ str(i) + "*J" + str(i) + "/100000000"   #시가총액(억) = 상장주식수 * 현재가

		table_df = table_df_list[3] #주요 재무제표
		table_df.columns = table_df.columns.droplevel(2)
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 23).Value = df.loc[0][1] # 2017.12 (Y) 매출액
		mWorkSheet.Cells(i, 24).Value = df.loc[0][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 25).Value = df.loc[0][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 26).Value = df.loc[0][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 28).Value = df.loc[0][5] # 2019.06
		mWorkSheet.Cells(i, 29).Value = df.loc[0][6] # 2019.09
		mWorkSheet.Cells(i, 30).Value = df.loc[0][7] # 2019.12
		mWorkSheet.Cells(i, 31).Value = df.loc[0][8] # 2020.03
		mWorkSheet.Cells(i, 32).Value = df.loc[0][9] # 2020.06
		mWorkSheet.Cells(i, 33).Value = df.loc[0][10] # 2020.09 (E)
		#mWorkSheet.Cells(i, 34).Value = df.loc[0][11] # 2020.12 (E)

		mWorkSheet.Cells(i, 35).Value = df.loc[1][1] # 2017.12 (Y) 영업이익
		mWorkSheet.Cells(i, 36).Value = df.loc[1][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 37).Value = df.loc[1][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 38).Value = df.loc[1][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 40).Value = df.loc[1][5] # 2019.06
		mWorkSheet.Cells(i, 41).Value = df.loc[1][6] # 2019.09
		mWorkSheet.Cells(i, 42).Value = df.loc[1][7] # 2019.12
		mWorkSheet.Cells(i, 43).Value = df.loc[1][8] # 2020.03
		mWorkSheet.Cells(i, 44).Value = df.loc[1][9] # 2020.06
		mWorkSheet.Cells(i, 45).Value = df.loc[1][10] # 2020.09 (E)
		#mWorkSheet.Cells(i, 46).Value = df.loc[1][11] # 2020.12 (E)

		mWorkSheet.Cells(i, 47).Value = df.loc[2][1] # 2017.12 (Y) 당기순이익
		mWorkSheet.Cells(i, 48).Value = df.loc[2][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 49).Value = df.loc[2][3] # 2019.12 (Y)
		mWorkSheet.Cells(i, 50).Value = df.loc[2][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 52).Value = df.loc[2][5] # 2019.06
		mWorkSheet.Cells(i, 53).Value = df.loc[2][6] # 2019.09
		mWorkSheet.Cells(i, 54).Value = df.loc[2][7] # 2019.12
		mWorkSheet.Cells(i, 55).Value = df.loc[2][8] # 2020.03
		mWorkSheet.Cells(i, 56).Value = df.loc[2][9] # 2020.06
		mWorkSheet.Cells(i, 57).Value = df.loc[2][10] # 2020.09 (E)
		#mWorkSheet.Cells(i, 58).Value = df.loc[2][10] # 2020.12 (E)

		mWorkSheet.Cells(i, 13).Value = df.loc[10][4] # PER = 주가 / 주당순이익(EPS)
		vTempEPS = df.loc[9][10] 	# 2020.09 EPS
		vTempEPS = vTempEPS*4
		mWorkSheet.Cells(i, 14).Value = int(mWorkSheet.Cells(i, 8).Value) / vTempEPS

		mWorkSheet.Cells(i, 15).Value = df.loc[12][4] # PBR

		mWorkSheet.Cells(i, 16).Value = df.loc[5][4] # ROE

		mWorkSheet.Cells(i, 59).Value = "=AI" + str(i) + "/W" + str(i)   #영업이익률(2017.12)
		mWorkSheet.Cells(i, 60).Value = "=AJ" + str(i) + "/X" + str(i)   #영업이익률(2018.12)
		mWorkSheet.Cells(i, 61).Value = "=AK" + str(i) + "/Y" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 62).Value = "=AL" + str(i) + "/Z" + str(i)   #영업이익률(2020.12)

		mWorkSheet.Cells(i, 64).Value = "=AN" +str(i) + "/AB" + str(i)   #영업이익률(2019.06)
		mWorkSheet.Cells(i, 65).Value = "=AO" +str(i) + "/AC" + str(i)   #영업이익률(2019.09)
		mWorkSheet.Cells(i, 66).Value = "=AP" +str(i) + "/AD" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 67).Value = "=AQ" +str(i) + "/AE" + str(i)   #영업이익률(2020.03)
		mWorkSheet.Cells(i, 68).Value = "=AR" +str(i) + "/AF" + str(i)   #영업이익률(2020.06)
		mWorkSheet.Cells(i, 69).Value = "=AS" +str(i) + "/AG" + str(i)  #영업이익률(2020.09)
		#mWorkSheet.Cells(i, 70).Value = "=AT" +str(i) + "/AH" + str(i)  #영업이익률(2020.12) (E)

		mWorkSheet.Cells(i, 71).Value = "=AU" +str(i) + "/W" + str(i)   #순이익률(2017.12)
		mWorkSheet.Cells(i, 72).Value = "=AV" +str(i) + "/X" + str(i)   #순이익률(2018.12)
		mWorkSheet.Cells(i, 73).Value = "=AW" +str(i) + "/Y" + str(i)   #순이익률(2019.12)
		mWorkSheet.Cells(i, 74).Value = "=AX" +str(i) + "/Z" + str(i)   #순이익률(2020.12)

		mWorkSheet.Cells(i, 76).Value = "=AZ" +str(i) + "/AB" + str(i)   #순이익률(2019.06)
		mWorkSheet.Cells(i, 77).Value = "=BA" +str(i) + "/AC" + str(i)   #순이익률(2019.09)
		mWorkSheet.Cells(i, 78).Value = "=BB" +str(i) + "/AD" + str(i)   #순이익률(2019.12)
		mWorkSheet.Cells(i, 79).Value = "=BC" +str(i) + "/AE" + str(i)   #순이익률(2020.03)
		mWorkSheet.Cells(i, 80).Value = "=BD" +str(i) + "/AF" + str(i)   #순이익률(2020.06)
		mWorkSheet.Cells(i, 81).Value = "=BE" +str(i) + "/AG" + str(i)  #순이익률(2020.09)
		#mWorkSheet.Cells(i, 82).Value = "=BF" +str(i) + "/AH" + str(i)  #순이익률(2020.12) (E)

		for j in range (1, 10) :
			if df.loc[5][j] == 65535 :
				mWorkSheet.Cells(i, 82+j).Value = 0
			else :
				mWorkSheet.Cells(i, 82+j).Value = df.loc[5][j] # ROE 2017.12 ~ 2020.12

		for j in range (1, 10) :
			if df.loc[6][j] == 65535 :
				mWorkSheet.Cells(i, 93+j).Value = 0
			else :
				mWorkSheet.Cells(i, 93+j).Value = df.loc[6][j] # 부채비율 2017.12 ~ 2020.12

		table_df = table_df_list[4]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 12).Value = df.loc[4][1] #외국인비율(%)

		mWorkSheet.Cells(i, 17).Value = "=Z" +str(i) + "/Y" + str(i)    # 2020 / 2019 매출증가 (YoY)
		mWorkSheet.Cells(i, 18).Value = "=AG" +str(i) + "/AC" + str(i)  # 전년동분기대비 매출증가 (QoQ)
		mWorkSheet.Cells(i, 19).Value = "=AL" +str(i) + "/AK" + str(i)  # 2020 / 2019 영업이익증가  (YoY)
		mWorkSheet.Cells(i, 20).Value = "=AS" +str(i) + "/AR" + str(i)  # 전년동분기대비 영업이익 증가  (QoQ)
		mWorkSheet.Cells(i, 21).Value = "=AX" +str(i) + "/AW" + str(i)  # 2020 / 2019 당기순이익증가 (YoY)
		mWorkSheet.Cells(i, 22).Value = "=BE" +str(i) + "/BA" + str(i)  # 전년동분기대비 당기순이익증가 (QoQ)

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
#table_df = table_df_list[3]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
#table_df.columns = table_df.columns.droplevel(2)
#print('table_df_list[3]')
#print(table_df)

mWorkBook.Save()
mWorkBook.Close()