#!/usr/bin/env python

import requests 
import pandas as pd
import sys
import re
import os
import win32com.client as win32

def set_title_list() :
	mWorkSheet.Cells(1, 1).Value = "관심종목코드"
	mWorkSheet.Cells(1, 2).Value = "관심종목명"
	mWorkSheet.Cells(1, 3).Value = "코드"
	mWorkSheet.Cells(1, 4).Value = "종목명"
	mWorkSheet.Cells(1, 5).Value = "소속"
	mWorkSheet.Cells(1, 6).Value = "섹터"
	mWorkSheet.Cells(1, 7).Value = "업종"
	mWorkSheet.Cells(1, 8).Value = "그룹"
	mWorkSheet.Cells(1, 9).Value = "테마"
	mWorkSheet.Cells(1, 10).Value = "종가"
	mWorkSheet.Cells(1, 11).Value = "등락금액 (%)"
	mWorkSheet.Cells(1, 12).Value = "상장주식수"
	mWorkSheet.Cells(1, 13).Value = "시총(억)"
	mWorkSheet.Cells(1, 14).Value = "외국인비율(%)"
	mWorkSheet.Cells(1, 15).Value = "PER (2020)"
	mWorkSheet.Cells(1, 16).Value = "PER (3분기*4)"
	mWorkSheet.Cells(1, 17).Value = "PBR (2020)"
	mWorkSheet.Cells(1, 18).Value = "ROE (지배주주)"
	mWorkSheet.Cells(1, 19).Value = "매출증가(YoY)"
	mWorkSheet.Cells(1, 20).Value = "매출증가 (전년동분기대비)"
	mWorkSheet.Cells(1, 21).Value = "영업이익증가 (YoY)"
	mWorkSheet.Cells(1, 22).Value = "영업이익증가 (전년동분기대비)"
	mWorkSheet.Cells(1, 23).Value = "당순증가(YoY)"
	mWorkSheet.Cells(1, 24).Value = "당순증가(전년동분기대비)"
	mWorkSheet.Cells(1, 25).Value = "매출액2017.12"
	mWorkSheet.Cells(1, 26).Value = "매출액2018.12"
	mWorkSheet.Cells(1, 27).Value = "매출액2019.12"
	mWorkSheet.Cells(1, 28).Value = "매출액2020.12(E)"
	mWorkSheet.Cells(1, 29).Value = "매출액2019.03"
	mWorkSheet.Cells(1, 30).Value = "매출액2019.06"
	mWorkSheet.Cells(1, 31).Value = "매출액2019.09"
	mWorkSheet.Cells(1, 32).Value = "매출액2019.12"
	mWorkSheet.Cells(1, 33).Value = "매출액2020.03"
	mWorkSheet.Cells(1, 34).Value = "매출액2020.06"
	mWorkSheet.Cells(1, 35).Value = "매출액2020.09"
	mWorkSheet.Cells(1, 36).Value = "매출액2020.12 (E)"
	mWorkSheet.Cells(1, 37).Value = "영업이익2017.12"
	mWorkSheet.Cells(1, 38).Value = "영업이익2018.12"
	mWorkSheet.Cells(1, 39).Value = "영업이익2019.12"
	mWorkSheet.Cells(1, 40).Value = "영업이익2020.12(E)"
	mWorkSheet.Cells(1, 41).Value = "영업이익2019.03"
	mWorkSheet.Cells(1, 42).Value = "영업이익2019.06"
	mWorkSheet.Cells(1, 43).Value = "영업이익2019.09"
	mWorkSheet.Cells(1, 44).Value = "영업이익2019.12"
	mWorkSheet.Cells(1, 45).Value = "영업이익2020.03"
	mWorkSheet.Cells(1, 46).Value = "영업이익2020.06"
	mWorkSheet.Cells(1, 47).Value = "영업이익2020.09"
	mWorkSheet.Cells(1, 48).Value = "영업이익2020.12(E)"
	mWorkSheet.Cells(1, 49).Value = "당기순이익2017.12"
	mWorkSheet.Cells(1, 50).Value = "당기순이익2018.12"
	mWorkSheet.Cells(1, 51).Value = "당기순이익2019.12"
	mWorkSheet.Cells(1, 52).Value = "당기순이익2020.12(E)"
	mWorkSheet.Cells(1, 53).Value = "당기순이익2019.03"
	mWorkSheet.Cells(1, 54).Value = "당기순이익2019.06"
	mWorkSheet.Cells(1, 55).Value = "당기순이익2019.09"
	mWorkSheet.Cells(1, 56).Value = "당기순이익2019.12"
	mWorkSheet.Cells(1, 57).Value = "당기순이익2020.03"
	mWorkSheet.Cells(1, 58).Value = "당기순이익2020.06"
	mWorkSheet.Cells(1, 59).Value = "당기순이익2020.09"
	mWorkSheet.Cells(1, 60).Value = "당기순이익2020.12(E)"
	mWorkSheet.Cells(1, 61).Value = "영업이익률2017.12"
	mWorkSheet.Cells(1, 62).Value = "영업이익률2018.12"
	mWorkSheet.Cells(1, 63).Value = "영업이익률2019.12"
	mWorkSheet.Cells(1, 64).Value = "영업이익률2020.12(E)"
	mWorkSheet.Cells(1, 65).Value = "영업이익률2019.03"
	mWorkSheet.Cells(1, 66).Value = "영업이익률2019.06"
	mWorkSheet.Cells(1, 67).Value = "영업이익률2019.09"
	mWorkSheet.Cells(1, 68).Value = "영업이익률2019.12"
	mWorkSheet.Cells(1, 69).Value = "영업이익률2020.03"
	mWorkSheet.Cells(1, 70).Value = "영업이익률2020.06"
	mWorkSheet.Cells(1, 71).Value = "영업이익률2020.09"
	mWorkSheet.Cells(1, 72).Value = "영업이익률2020.12(E)"
	mWorkSheet.Cells(1, 73).Value = "순이익률2017.12"
	mWorkSheet.Cells(1, 74).Value = "순이익률2018.12"
	mWorkSheet.Cells(1, 75).Value = "순이익률2019.12"
	mWorkSheet.Cells(1, 76).Value = "순이익률2020.12(E)"
	mWorkSheet.Cells(1, 77).Value = "순이익률2019.03"
	mWorkSheet.Cells(1, 78).Value = "순이익률2019.06"
	mWorkSheet.Cells(1, 79).Value = "순이익률2019.09"
	mWorkSheet.Cells(1, 80).Value = "순이익률2019.12"
	mWorkSheet.Cells(1, 81).Value = "순이익률2020.03"
	mWorkSheet.Cells(1, 82).Value = "순이익률2020.06"
	mWorkSheet.Cells(1, 83).Value = "순이익률2020.09"
	mWorkSheet.Cells(1, 84).Value = "순이익률2020.12(E)"
	mWorkSheet.Cells(1, 85).Value = "ROE2017.12"
	mWorkSheet.Cells(1, 86).Value = "ROE2018.12"
	mWorkSheet.Cells(1, 87).Value = "ROE2019.12"
	mWorkSheet.Cells(1, 88).Value = "ROE2020.12(E)"
	mWorkSheet.Cells(1, 89).Value = "ROE2019.06"
	mWorkSheet.Cells(1, 90).Value = "ROE2019.09"
	mWorkSheet.Cells(1, 91).Value = "ROE2019.12"
	mWorkSheet.Cells(1, 92).Value = "ROE2020.03"
	mWorkSheet.Cells(1, 93).Value = "ROE2020.06"
	mWorkSheet.Cells(1, 94).Value = "ROE2020.09"
	mWorkSheet.Cells(1, 95).Value = "ROE2020.12(E)"
	mWorkSheet.Cells(1, 96).Value = "부채비율2017.12"
	mWorkSheet.Cells(1, 97).Value = "부채비율2018.12"
	mWorkSheet.Cells(1, 98).Value = "부채비율2019.12"
	mWorkSheet.Cells(1, 99).Value = "부채비율2020.12(E)"
	mWorkSheet.Cells(1, 100).Value = "부채비율2019.06"
	mWorkSheet.Cells(1, 101).Value = "부채비율2019.09"
	mWorkSheet.Cells(1, 102).Value = "부채비율2019.12"
	mWorkSheet.Cells(1, 103).Value = "부채비율2020.03"
	mWorkSheet.Cells(1, 104).Value = "부채비율2020.06"
	mWorkSheet.Cells(1, 105).Value = "부채비율2020.09"
	mWorkSheet.Cells(1, 106).Value = "부채비율2020.12(E)"

def get_company_list() :
#copy csv data to xlsx
	lastCol = mWorkSheet.UsedRange.Columns.Count
	lastRow = mWorkSheet.UsedRange.Rows.Count
	print("Laster Row : " + str(lastRow) + " colum num : " + str(lastCol))
	#mWorkSheet.Range("C2:C" + str(lastRow)).Select()  # 기업코드

	for i in range (2270, lastRow+1) :
		url = 'https://finance.naver.com/item/sise.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)
		print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함
		table_df = table_df_list[1]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
		df = pd.DataFrame(table_df)
		# df.loc = row, column, 0부터 시작
		mWorkSheet.Cells(i, 10).Value = df.loc[0][1] #현재가
		mWorkSheet.Cells(i, 11).Value = str(df.loc[1][1]) + " " + str(df.loc[2][1]) #전일대비 + 등락률(%)
		#print(df.head()) #print dataframe data
		#print(df.shape) #get row, column count

		url = 'https://finance.naver.com/item/main.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)
		#print('url = ' + str(url))
		table_df_list = pd.read_html(url, encoding='euc-kr')
		table_df = table_df_list[5] #주요 재무제표
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 12).Value = df.loc[2][1] #상장주식수
		mWorkSheet.Cells(i, 13).Value = "=J"+ str(i) + "*L" + str(i) + "/100000000"   #시가총액(억) = 상장주식수 * 현재가

		table_df = table_df_list[3] #주요 재무제표
		table_df.columns = table_df.columns.droplevel(2)
		df = pd.DataFrame(table_df)

		#매출액
		mWorkSheet.Cells(i, 25).Value = df.loc[0][1] # 2017.12 (Y)
		mWorkSheet.Cells(i, 26).Value = df.loc[0][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 27).Value = df.loc[0][3] # 2019.12 (Y)
		if pd.isnull(df.loc[0, 4]) == True :
			#어닝이 없을 경우 이전 2019년 3,4분기 + 2020년 1,2분기를 더한다. 
			vTmp = df.loc[0][6] +  df.loc[0][7] + df.loc[0][8] + df.loc[0][9]
			mWorkSheet.Cells(i, 28).Value = vTmp
		else :
			mWorkSheet.Cells(i, 28).Value = df.loc[0][4] # 2020.12 (E) (Y)		

		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 30).Value = df.loc[0][5] # 2019.06
		mWorkSheet.Cells(i, 31).Value = df.loc[0][6] # 2019.09
		mWorkSheet.Cells(i, 32).Value = df.loc[0][7] # 2019.12
		mWorkSheet.Cells(i, 33).Value = df.loc[0][8] # 2020.03
		mWorkSheet.Cells(i, 34).Value = df.loc[0][9] # 2020.06
		if pd.isnull(df.loc[0, 10]) == True :
			# 어닝이 없을 경우 이전 2019년 3분기를 사용한다. 
			mWorkSheet.Cells(i, 35).Value = df.loc[0][6] # 2019.09
		else :
			mWorkSheet.Cells(i, 35).Value = df.loc[0][10] # 2020.09 (E)
		#mWorkSheet.Cells(i, 36).Value = df.loc[0][11] # 2020.12 (E)

		#영업이익
		mWorkSheet.Cells(i, 37).Value = df.loc[1][1] # 2017.12 (Y)
		mWorkSheet.Cells(i, 38).Value = df.loc[1][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 39).Value = df.loc[1][3] # 2019.12 (Y)
		if pd.isnull(df.loc[1, 4]) == True :
			#어닝이 없을 경우 이전 2019년 3,4분기 + 2020년 1,2분기를 더한다. 
			vTmp = df.loc[1][6] +  df.loc[1][7] + df.loc[1][8] + df.loc[1][9]
			mWorkSheet.Cells(i, 40).Value = vTmp
		else :
			mWorkSheet.Cells(i, 40).Value = df.loc[1][4] # 2020.12 (E) (Y)

		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 42).Value = df.loc[1][5] # 2019.06
		mWorkSheet.Cells(i, 43).Value = df.loc[1][6] # 2019.09
		mWorkSheet.Cells(i, 44).Value = df.loc[1][7] # 2019.12
		mWorkSheet.Cells(i, 45).Value = df.loc[1][8] # 2020.03
		mWorkSheet.Cells(i, 46).Value = df.loc[1][9] # 2020.06
		if pd.isnull(df.loc[1, 10]) == True :
			# 어닝이 없을 경우 이전 2019년 3분기를 사용한다. 
			mWorkSheet.Cells(i, 47).Value = df.loc[1][6] # 2019.09
		else :
			mWorkSheet.Cells(i, 47).Value = df.loc[1][10] # 2020.12 (E) (Y)
		#mWorkSheet.Cells(i, 48).Value = df.loc[1][11] # 2020.12 (E)

		#당기순이익
		mWorkSheet.Cells(i, 49).Value = df.loc[2][1] # 2017.12 (Y)
		mWorkSheet.Cells(i, 50).Value = df.loc[2][2] # 2018.12 (Y)
		mWorkSheet.Cells(i, 51).Value = df.loc[2][3] # 2019.12 (Y)
		if pd.isnull(df.loc[2, 4]) == True :
			#어닝이 없을 경우 이전 2019년 3,4분기 + 2020년 1,2분기를 더한다. 
			vTmp = df.loc[2][6] +  df.loc[2][7] + df.loc[2][8] + df.loc[2][9]
			mWorkSheet.Cells(i, 52).Value = vTmp
		else :
			mWorkSheet.Cells(i, 52).Value = df.loc[2][4] # 2020.12 (E) (Y)
		#to be filled 2020.03 data
		mWorkSheet.Cells(i, 54).Value = df.loc[2][5] # 2019.06
		mWorkSheet.Cells(i, 55).Value = df.loc[2][6] # 2019.09
		mWorkSheet.Cells(i, 56).Value = df.loc[2][7] # 2019.12
		mWorkSheet.Cells(i, 57).Value = df.loc[2][8] # 2020.03
		mWorkSheet.Cells(i, 58).Value = df.loc[2][9] # 2020.06
		if pd.isnull(df.loc[2, 10]) == True :
			# 어닝이 없을 경우 이전 2019년 3분기를 사용한다. 
			mWorkSheet.Cells(i, 59).Value = df.loc[2][6] # 2019.09
		else :
			mWorkSheet.Cells(i, 59).Value = df.loc[2][10] # 2020.09 (E)
		#mWorkSheet.Cells(i, 60).Value = df.loc[2][10] # 2020.12 (E)

		mWorkSheet.Cells(i, 15).Value = df.loc[10][4]  # PER = 주가 / 주당순이익(EPS)
		vTempEPS = df.loc[9][10] 	                   # 2020.09 EPS
		vTempEPS = vTempEPS*4
		mWorkSheet.Cells(i, 16).Value = float(str(mWorkSheet.Cells(i, 10).Value)) / vTempEPS  # 주가 / 주당 순이익

		mWorkSheet.Cells(i, 17).Value = df.loc[12][4] # PBR

		mWorkSheet.Cells(i, 18).Value = df.loc[5][4] # ROE

		mWorkSheet.Cells(i, 61).Value = "=AK" + str(i) + "/Y" + str(i)   #영업이익률(2017.12)
		mWorkSheet.Cells(i, 62).Value = "=AL" + str(i) + "/Z" + str(i)   #영업이익률(2018.12)
		mWorkSheet.Cells(i, 63).Value = "=AM" + str(i) + "/AA" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 64).Value = "=AN" + str(i) + "/AB" + str(i)   #영업이익률(2020.12)

		mWorkSheet.Cells(i, 66).Value = "=AP" +str(i) + "/AD" + str(i)   #영업이익률(2019.06)
		mWorkSheet.Cells(i, 67).Value = "=AQ" +str(i) + "/AE" + str(i)   #영업이익률(2019.09)
		mWorkSheet.Cells(i, 68).Value = "=AR" +str(i) + "/AF" + str(i)   #영업이익률(2019.12)
		mWorkSheet.Cells(i, 69).Value = "=AS" +str(i) + "/AG" + str(i)   #영업이익률(2020.03)
		mWorkSheet.Cells(i, 70).Value = "=AT" +str(i) + "/AH" + str(i)   #영업이익률(2020.06)
		mWorkSheet.Cells(i, 71).Value = "=AU" +str(i) + "/AI" + str(i)  #영업이익률(2020.09)
		#mWorkSheet.Cells(i, 72).Value = "=AV" +str(i) + "/AJ" + str(i)  #영업이익률(2020.12) (E)

		mWorkSheet.Cells(i, 73).Value = "=AW" +str(i) + "/Y" + str(i)   #순이익률(2017.12)
		mWorkSheet.Cells(i, 74).Value = "=AX" +str(i) + "/Z" + str(i)   #순이익률(2018.12)
		mWorkSheet.Cells(i, 75).Value = "=AY" +str(i) + "/AA" + str(i)   #순이익률(2019.12)
		mWorkSheet.Cells(i, 76).Value = "=AZ" +str(i) + "/AB" + str(i)   #순이익률(2020.12)

		mWorkSheet.Cells(i, 78).Value = "=BB" +str(i) + "/AD" + str(i)   #순이익률(2019.06)
		mWorkSheet.Cells(i, 79).Value = "=BC" +str(i) + "/AE" + str(i)   #순이익률(2019.09)
		mWorkSheet.Cells(i, 80).Value = "=BD" +str(i) + "/AF" + str(i)   #순이익률(2019.12)
		mWorkSheet.Cells(i, 81).Value = "=BE" +str(i) + "/AG" + str(i)   #순이익률(2020.03)
		mWorkSheet.Cells(i, 82).Value = "=BF" +str(i) + "/AH" + str(i)   #순이익률(2020.06)
		mWorkSheet.Cells(i, 83).Value = "=BG" +str(i) + "/AI" + str(i)  #순이익률(2020.09)
		#mWorkSheet.Cells(i, 84).Value = "=BH" +str(i) + "/AJ" + str(i)  #순이익률(2020.12) (E)

		for j in range (1, 10) :
			if pd.isnull(df.loc[5, j]) == True : # Is NaN
				mWorkSheet.Cells(i, 84+j).Value = 0
			else :
				mWorkSheet.Cells(i, 84+j).Value = df.loc[5][j] # ROE 2017.12 ~ 2020.12

		for j in range (1, 10) :
			if pd.isnull(df.loc[6, j]) == True : # Is NaN
				mWorkSheet.Cells(i, 95+j).Value = 0
			else :
				mWorkSheet.Cells(i, 95+j).Value = df.loc[6][j] # 부채비율 2017.12 ~ 2020.12

		table_df = table_df_list[4]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 14).Value = df.loc[4][1] #외국인비율(%)

		mWorkSheet.Cells(i, 19).Value = "=AB" +str(i) + "/AA" + str(i)    # 2020 / 2019 매출증가 (YoY)
		mWorkSheet.Cells(i, 20).Value = "=AI" +str(i) + "/AE" + str(i)  # 전년동분기대비 매출증가 (QoQ)
		mWorkSheet.Cells(i, 21).Value = "=AN" +str(i) + "/AM" + str(i)  # 2020 / 2019 영업이익증가  (YoY)
		mWorkSheet.Cells(i, 22).Value = "=AU" +str(i) + "/AQ" + str(i)  # 전년동분기대비 영업이익 증가  (QoQ)
		mWorkSheet.Cells(i, 23).Value = "=AZ" +str(i) + "/AY" + str(i)  # 2020 / 2019 당기순이익증가 (YoY)
		mWorkSheet.Cells(i, 24).Value = "=BG" +str(i) + "/BC" + str(i)  # 전년동분기대비 당기순이익증가 (QoQ)

filename = "export_Data.xlsx"
filepath = os.path.abspath(filename)
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

mWorkBook = excel.Workbooks.Open(filepath)
mWorkSheet = mWorkBook.Worksheets('RawData')
mWorkSheet.Select()

#set_title_list()
get_company_list()
#def run_each_company_data(company_code) :

#test code start
#url = 'https://finance.naver.com/item/main.nhn?code=071840'
#table_df_list = pd.read_html(url, encoding='euc-kr')
#table_df = table_df_list[0]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
#table_df.columns = table_df.columns.droplevel(2)
#print('table_df_list[0]')
#print(table_df)
#test code end

mWorkBook.Save()
mWorkBook.Close()
