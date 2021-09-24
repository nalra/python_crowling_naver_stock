#!/usr/bin/env python

import requests 
import pandas as pd
import numpy as np
import sys
import re
import os
import win32com.client as win32

#pip install requests_html
#pip install numpy
#pip install pandas
#pip install pypiwin32 - win32com

def get_company_list_full() :
	lastCol = mWorkSheet.UsedRange.Columns.Count
	lastRow = mWorkSheet.UsedRange.Rows.Count
	print("Laster Row : " + str(lastRow) + " colum num : " + str(lastCol))
	#mWorkSheet.Range("C2:C" + str(lastRow)).Select()  # 기업코드
	# refer https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference

	for i in range (2, lastRow) :
		url = 'https://finance.naver.com/item/sise.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)
		try :
			table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함
		except :
			print('NULL url = ' + str(url))
			continue
		print('url = ' + str(url))
		table_df = table_df_list[1]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
		df = pd.DataFrame(table_df)
		# df.iloc = row, column, 0부터 시작
		mWorkSheet.Cells(i, 10).Value = df.iloc[0][1] #현재가
		mWorkSheet.Cells(i, 11).Value = str(df.iloc[1][1]) + " " + str(df.iloc[2][1]) #전일대비 + 등락률(%)
		#print(df.head()) #print dataframe data
		#print(df.shape) #get row, column count

		url = 'https://finance.naver.com/item/main.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)

		table_df_list = pd.read_html(url, encoding='euc-kr')
		table_df = table_df_list[5] #주요 재무제표
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 12).Value = df.iloc[2][1] #상장주식수
		mWorkSheet.Cells(i, 13).Value = "=J"+ str(i) + "*L" + str(i) + "/100000000"   #시가총액(억) = 상장주식수 * 현재가

		table_df = table_df_list[7]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 141).Value = df.iloc[1][0] #52주 최고
		mWorkSheet.Cells(i, 142).Value = df.iloc[1][1] #52주 최저

		if mWorkSheet.Cells(i, 9).Value == str('SPAC') or mWorkSheet.Cells(i, 9).Value == str('리츠') :
			continue

		table_df = table_df_list[3] #주요 재무제표
		#table_df.columns = table_df.columns.droplevel(2)
		df = pd.DataFrame(table_df)
		df.fillna(0)

		#매출액
		mWorkSheet.Cells(i, 26).Value = df.iloc[0][1] # 2018.12 (Y)
		mWorkSheet.Cells(i, 27).Value = df.iloc[0][2] # 2019.12 (Y)
		mWorkSheet.Cells(i, 28).Value = df.iloc[0][3] # 2020.12 (Y)
		if pd.isna(df.iloc[0,4]) :
			#어닝이 없을 경우 최근 4분기를 더한다. 
			vSumRevenue = 0
			if pd.isna(df.iloc[0,6]) or (str(df.iloc[0,6]) == str('-')) :
				vSumRevenue += 0
			else : 
				vSumRevenue += float(df.iloc[0][6])
			if pd.isna(df.iloc[0,7]) or (str(df.iloc[0,7]) == str('-')) :
				vSumRevenue += 0
			else : 
				vSumRevenue += float(df.iloc[0][7])
			if pd.isna(df.iloc[0,8]) or (str(df.iloc[0,8]) == str('-')) :
				vSumRevenue += 0
			else :
				vSumRevenue += float(df.iloc[0][8])
			if pd.isna(df.iloc[0,9]) or (str(df.iloc[0,9]) == str('-')) :
				vSumRevenue += 0
			else :
				vSumRevenue = float(df.iloc[0][9])
			mWorkSheet.Cells(i, 29).Value = vSumRevenue
		else :
			mWorkSheet.Cells(i, 29).Value = df.iloc[0][4] # 2021.12 (E) (Y)
		#print("매출액 : " + str(mWorkSheet.Cells(i, 28).Value))

		#to be filled from 2020.06 data
		mWorkSheet.Cells(i, 35).Value = df.iloc[0][5] # 2020.06     AI column
		mWorkSheet.Cells(i, 36).Value = df.iloc[0][6] # 2020.09
		mWorkSheet.Cells(i, 37).Value = df.iloc[0][7] # 2020.12
		mWorkSheet.Cells(i, 38).Value = df.iloc[0][8] # 2021.03
		mWorkSheet.Cells(i, 39).Value = df.iloc[0][9] # 2021.06
		if pd.isna(df.iloc[0,9]) or (str(df.iloc[0,9]) == str('-')) :
			mWorkSheet.Cells(i, 39).Value = 0
		else :
			mWorkSheet.Cells(i, 39).Value = df.iloc[0][9] # 2021.06

		if pd.isna(df.iloc[0,10]) or (str(df.iloc[0,10]) == str('-')) :
			# 어닝이 없을 경우 이전 2020년 3분기를 사용한다.
			if pd.isna(df.iloc[0,6]) or (str(df.iloc[0,6]) == str('-')) :
				mWorkSheet.Cells(i, 40).Value = 0
			else :
				mWorkSheet.Cells(i, 40).Value = int(df.iloc[0][6]) # 2020.09
		else :
			mWorkSheet.Cells(i, 40).Value = df.iloc[0][10] # 2021.09 (E)

		#영업이익
		mWorkSheet.Cells(i, 42).Value = df.iloc[1][1] # 2018.12 (Y)   AP column
		mWorkSheet.Cells(i, 43).Value = df.iloc[1][2] # 2019.12 (Y)
		mWorkSheet.Cells(i, 44).Value = df.iloc[1][3] # 2020.12 (Y)
		if pd.isna(df.iloc[1,4]) :
			#어닝이 없을 경우 최근 4분기를 더한다. 
			vSumProfit = 0
			if pd.isna(df.iloc[1,6]) or (str(df.iloc[1,6]) == str('-')) :
				vSumProfit += 0
			else :
				vSumProfit += float(df.iloc[1][6])
			if pd.isna(df.iloc[1,7]) or (str(df.iloc[1,7]) == str('-')) :
				vSumProfit += 0
			else :
				vSumProfit += float(df.iloc[1][7])
			if pd.isna(df.iloc[1,8]) or (str(df.iloc[1,8]) == str('-')) :
				vSumProfit += 0
			else :
				vSumProfit += float(df.iloc[1][8])
			if pd.isna(df.iloc[1,9]) or (str(df.iloc[1,9]) == str('-')) :
				vSumProfit += 0
			else :
				vSumProfit = float(df.iloc[1][9])
			mWorkSheet.Cells(i, 45).Value = vSumProfit
		else :
			mWorkSheet.Cells(i, 45).Value = df.iloc[1][4] # 2021.12 (E) (Y)

		#to be filled 2020.06 data
		mWorkSheet.Cells(i, 51).Value = df.iloc[1][5] # 2020.06       #AY column
		mWorkSheet.Cells(i, 52).Value = df.iloc[1][6] # 2020.09
		mWorkSheet.Cells(i, 53).Value = df.iloc[1][7] # 2020.12
		mWorkSheet.Cells(i, 54).Value = df.iloc[1][8] # 2021.03
		if pd.isna(df.iloc[1,9]) :
			mWorkSheet.Cells(i, 55).Value = 0
		else :
			mWorkSheet.Cells(i, 55).Value = df.iloc[1][9] # 2021.06
		if pd.isna(df.iloc[1,10]) or (str(df.iloc[1,10]) == str('-')) :
			# 어닝이 없을 경우 이전 2020년 3분기를 사용한다.
			if pd.isna(df.iloc[1,6]) or (str(df.iloc[1,6]) == str('-')) :
				mWorkSheet.Cells(i, 56).Value = 0
			else :
				mWorkSheet.Cells(i, 56).Value = int(df.iloc[1][6]) # 2020.09
		else :
			mWorkSheet.Cells(i, 56).Value = df.iloc[1][10] # 2021.09 (E)


		#당기순이익
		mWorkSheet.Cells(i, 58).Value = df.iloc[2][1] # 2018.12 (Y)
		mWorkSheet.Cells(i, 59).Value = df.iloc[2][2] # 2019.12 (Y)
		mWorkSheet.Cells(i, 60).Value = df.iloc[2][3] # 2020.12 (Y)
		if pd.isna(df.iloc[2,4]) :
			#어닝이 없을 경우 최근 4분기를 더한다. 
			vSumEarning = 0
			if pd.isna(df.iloc[2,6]) or (str(df.iloc[2,6]) == str('-')) :
				vSumEarning += 0
			else : 
				vSumEarning += float(df.iloc[2][6])
			if pd.isna(df.iloc[2,7]) or (str(df.iloc[2,7]) == str('-')) :
				vSumEarning += 0
			else : 
				vSumEarning += float(df.iloc[2][7])
			if pd.isna(df.iloc[2,8]) or (str(df.iloc[2,8]) == str('-')) :
				vSumEarning += 0
			else :
				vSumEarning += float(df.iloc[2][8])
			if pd.isna(df.iloc[2,9]) or (str(df.iloc[2,9]) == str('-')) :
				vSumEarning += 0
			else :
				vSumEarning = float(df.iloc[2][9])
			mWorkSheet.Cells(i, 61).Value = vSumEarning
		else :
			mWorkSheet.Cells(i, 61).Value = df.iloc[2][4] # 2021.12 (E) (Y)

		#to be filled 2021.06 data
		mWorkSheet.Cells(i, 67).Value = df.iloc[2][5] # 2020.06    BO Column
		mWorkSheet.Cells(i, 68).Value = df.iloc[2][6] # 2020.09
		mWorkSheet.Cells(i, 69).Value = df.iloc[2][7] # 2020.12
		mWorkSheet.Cells(i, 70).Value = df.iloc[2][8] # 2021.03
		if pd.isna(df.iloc[2,9]) :
			mWorkSheet.Cells(i, 71).Value = 0
		else :
			mWorkSheet.Cells(i, 71).Value = df.iloc[2][9] # 2021.06
		if pd.isna(df.iloc[2,10]) :
			# 어닝이 없을 경우 이전 2020년 3분기를 사용한다.
			mWorkSheet.Cells(i, 72).Value = mWorkSheet.Cells(i, 68).Value
		else :
			mWorkSheet.Cells(i, 72).Value = df.iloc[2][10] # 2021.09 (E)

		#EPS
		mWorkSheet.Cells(i, 98).Value = df.iloc[9][1] # 2018.12 (Y)
		mWorkSheet.Cells(i, 99).Value = df.iloc[9][2] # 2019.12 (Y)
		mWorkSheet.Cells(i, 100).Value = df.iloc[9][3] # 2020.12 (Y)
		if pd.isna(df.iloc[9,4]) :
			#어닝이 없을 경우 최근 4분기를 더한다. 
			vSumEarning = 0
			if pd.isna(df.iloc[9,6]) or (str(df.iloc[9,6]) == str('-')) :
				vSumEarning += 0
			else : 
				vSumEarning += float(df.iloc[9][6])
			if pd.isna(df.iloc[9,7]) or (str(df.iloc[9,7]) == str('-')) :
				vSumEarning += 0
			else : 
				vSumEarning += float(df.iloc[9][7])
			if pd.isna(df.iloc[9,8]) or (str(df.iloc[9,8]) == str('-')) :
				vSumEarning += 0
			else :
				vSumEarning += float(df.iloc[9][8])
			if pd.isna(df.iloc[9,9]) or (str(df.iloc[9,9]) == str('-')) :
				vSumEarning += 0
			else :
				vSumEarning = float(df.iloc[9][9])
			mWorkSheet.Cells(i, 101).Value = vSumEarning
		else :
			mWorkSheet.Cells(i, 101).Value = df.iloc[9][4] # 2021.12 (E) (Y)

		#to be filled 2021.06 data
		mWorkSheet.Cells(i, 102).Value = df.iloc[9][5] # 2020.06    CX Column
		mWorkSheet.Cells(i, 103).Value = df.iloc[9][6] # 2020.09
		mWorkSheet.Cells(i, 104).Value = df.iloc[9][7] # 2020.12
		mWorkSheet.Cells(i, 105).Value = df.iloc[9][8] # 2021.03
		if pd.isna(df.iloc[9,9]) :
			mWorkSheet.Cells(i, 106).Value = 0
		else :
			mWorkSheet.Cells(i, 106).Value = df.iloc[9][9] # 2021.06
		if pd.isna(df.iloc[9,10]) :
			# 어닝이 없을 경우 이전 2020년 3분기를 사용한다.
			mWorkSheet.Cells(i, 107).Value = mWorkSheet.Cells(i, 68).Value
		else :
			mWorkSheet.Cells(i, 107).Value = df.iloc[9][10] # 2021.09 (E)


		if pd.isna(df.iloc[10,4]) or (str(df.iloc[10,4]) == str('-')) :
			#PER 어닝 이 없는 경우 직접 PER을 구한다.
			vTempTotalEPS = 0
			vTmpQtrCnt = 0
			if pd.isna(df.iloc[9,6]) or (str(df.iloc[9,6]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][6])
			if pd.isna(df.iloc[9,7]) or (str(df.iloc[9,7]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][7])
			if pd.isna(df.iloc[9,8]) or (str(df.iloc[9,8]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][8])
			if pd.isna(df.iloc[9,9]) or (str(df.iloc[9,9]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][9])
			print("vTempTotalEPS = " + str(vTempTotalEPS))
			if vTempTotalEPS != 0 :
				vTempTotalEPS = float(mWorkSheet.Cells(i, 10).Value / float(vTempTotalEPS))
				# PER = 주가 / 주당순이익(EPS)
			mWorkSheet.Cells(i, 15).Value = vTempTotalEPS
			print("estimated EPS : " + str(vTempTotalEPS))
		else :
			mWorkSheet.Cells(i, 15).Value =  df.iloc[10][4]

		if pd.isna(df.iloc[9,9]) or (str(df.iloc[9,9]) == str('-')) :
			vTempPER = 0
		else :
			vTempEPS = int(df.iloc[9][9])   #  최근 마지막 실적 2021.06 EPS
			vTempEPS = vTempEPS*4
			if vTempEPS != 0 :				
				vTempPER = float(mWorkSheet.Cells(i, 10).Value / float(vTempEPS))
			else :
				vTempPER = 0
		mWorkSheet.Cells(i, 16).Value = vTempPER  # 최근 분기 대비 PER
		print("estimated quarter PER : " + str(vTempPER))


		if pd.isna(df.iloc[12,4])  or (str(df.iloc[12,4]) == str('-')) :
			# PBR 어닝이 없을 경우 직전 4분기 평균을 사용한다.
			vTmpQtrCnt = 0
			vTempPBR = 0
			if pd.isna(df.iloc[12,6]) or (str(df.iloc[12,6]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][6])
			if pd.isna(df.iloc[12,7]) or (str(df.iloc[12,7]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][7])
			if pd.isna(df.iloc[12,8]) or (str(df.iloc[12,8]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][8])
			if pd.isna(df.iloc[12,9]) or (str(df.iloc[12,9]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][9])
			print("estimated vTempPBR : " + str(vTempPBR))
			if vTmpQtrCnt != 0 :
				vTempPBR = float(vTempPBR/float(vTmpQtrCnt))
			mWorkSheet.Cells(i, 17).Value = vTempPBR
		else :
			mWorkSheet.Cells(i, 17).Value = df.iloc[12][4] # PBR
		print("estimated PBR : " + str(mWorkSheet.Cells(i, 17).Value))

		#ROE
		mWorkSheet.Cells(i, 112).Value = df.iloc[5][1] # 2018.12 (Y)
		mWorkSheet.Cells(i, 113).Value = df.iloc[5][2] # 2019.12 (Y)
		mWorkSheet.Cells(i, 114).Value = df.iloc[5][3] # 2020.12 (Y)
		if pd.isna(df.iloc[5,4]) :
			# ROE 어닝이 없을 경우 직전 4분기 평균을 사용한다.
			vTmpQtrCnt = 0
			vTempROE = 0
			if pd.isna(df.iloc[5,6]) or (str(df.iloc[5,6]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][6])
			if pd.isna(df.iloc[5,7]) or (str(df.iloc[5,7]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][7])
			if pd.isna(df.iloc[5,8]) or (str(df.iloc[5,8]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][8])
			if pd.isna(df.iloc[5,9]) or (str(df.iloc[5,8]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][9])
			print("estimated vTempROE : " + str(vTempROE))
			if vTmpQtrCnt != 0 :
				vTempROE = float(vTempROE/float(vTmpQtrCnt))
			mWorkSheet.Cells(i, 18).Value = vTempROE
		else :
			mWorkSheet.Cells(i, 18).Value = df.iloc[5][4] # ROE
			print("estimated ROE : " + str(mWorkSheet.Cells(i, 18).Value))

		mWorkSheet.Cells(i, 120).Value = df.iloc[5][5] # 2020.06    DP Column
		mWorkSheet.Cells(i, 121).Value = df.iloc[5][6] # 2020.09
		mWorkSheet.Cells(i, 122).Value = df.iloc[5][7] # 2020.12
		mWorkSheet.Cells(i, 123).Value = df.iloc[5][8] # 2021.03
		if pd.isna(df.iloc[5,9]) :
			mWorkSheet.Cells(i, 124).Value = 0
		else :
			mWorkSheet.Cells(i, 124).Value = df.iloc[5][9] # 2021.06
		if pd.isna(df.iloc[5,10]) :
			# 어닝이 없을 경우 위의 계산한 ROE를 사용한다.
			mWorkSheet.Cells(i, 125).Value = mWorkSheet.Cells(i, 18).Value
		else :
			mWorkSheet.Cells(i, 125).Value = df.iloc[5][10] # 2021.09 (E)


		mWorkSheet.Cells(i, 86).Value = "=BB" +str(i) + "/AL" + str(i)   #영업이익률(2021.03)
		mWorkSheet.Cells(i, 87).Value = "=BC" +str(i) + "/AM" + str(i)   #영업이익률(2021.06)
		mWorkSheet.Cells(i, 88).Value = "=BD" +str(i) + "/AN" + str(i)   #영업이익률(2021.09) (E)

		mWorkSheet.Cells(i, 93).Value = "=BI" +str(i) + "/AC" + str(i)   #순이익률(2021.12)

		mWorkSheet.Cells(i, 94).Value = "=BQ" +str(i) + "/AK" + str(i)   #순이익률(2020.12)
		mWorkSheet.Cells(i, 95).Value = "=BR" +str(i) + "/AL" + str(i)   #순이익률(2021.03)
		mWorkSheet.Cells(i, 96).Value = "=BS" +str(i) + "/AM" + str(i)   #순이익률(2021.06)
		mWorkSheet.Cells(i, 97).Value = "=BT" +str(i) + "/AN" + str(i)   #순이익률(2021.09) (E)

		mWorkSheet.Cells(i, 108).Value = "=(CU" +str(i) + "/CT" + str(i) + ") -1"   #EPS성장률 (2018~2019)
		mWorkSheet.Cells(i, 109).Value = "=(CV" +str(i) + "/CU" + str(i) + ") -1"   #EPS성장률 (2019~2020)
		mWorkSheet.Cells(i, 110).Value = "=(CW" +str(i) + "/CV" + str(i) + ") -1"   #EPS성장률 (2020~2021)
		mWorkSheet.Cells(i, 145).Value = "=O" +str(i) + "/AVERAGE(DD" + str(i) + ":DF" + str(i) +")"   #PEG (PER / EPS성장률3년)
		mWorkSheet.Cells(i, 146).Value = "=O" +str(i) + "/AVERAGE(DE" + str(i) + ":DF" + str(i) +")"  #PEG (PER / EPS성장률2년)
		mWorkSheet.Cells(i, 147).Value = "=O" +str(i) + "/DF" + str(i)  #PEG (PER / EPS성장률1년)


		for j in range (1, 4) :
			if pd.isna(df.iloc[6,j]) or (str(df.iloc[6,j]) == str('-')) :
				mWorkSheet.Cells(i, 126+j).Value = 0
			else :
				mWorkSheet.Cells(i, 126+j).Value = df.iloc[6][j] # 부채비율 2018.12 ~ 2021.12
		for j in range (5, 10) :
			if pd.isna(df.iloc[6,j]) or (str(df.iloc[6,j]) == str('-')) :
				mWorkSheet.Cells(i, 130+j).Value = 0                       # EE Column
			else :
				mWorkSheet.Cells(i, 130+j).Value = df.iloc[6][j] # 부채비율 2020.06 ~ 2021.09

		table_df = table_df_list[4]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 14).Value = df.iloc[4][1] #외국인비율(%)

		mWorkSheet.Cells(i, 19).Value = "=AC" +str(i) + "/AB" + str(i)  # 2021 / 2020 매출증가 (YoY)
		mWorkSheet.Cells(i, 20).Value = "=AN" +str(i) + "/AJ" + str(i)  # 전년동분기대비 매출증가 (QoQ)
		mWorkSheet.Cells(i, 21).Value = "=AS" +str(i) + "/AR" + str(i)  # 2021 / 2020 영업이익증가  (YoY)
		mWorkSheet.Cells(i, 22).Value = "=BD" +str(i) + "/AZ" + str(i)  # 전년동분기대비 영업이익 증가  (QoQ)
		mWorkSheet.Cells(i, 23).Value = "=BI" +str(i) + "/BH" + str(i)  # 2021 / 2020 당기순이익증가 (YoY)
		mWorkSheet.Cells(i, 24).Value = "=BT" +str(i) + "/BP" + str(i)  # 전년동분기대비 당기순이익증가 (QoQ)

		if i % 100 == 0 :
			mWorkBook.Save() #work around for excel fault


def get_company_list_value() :
	lastCol = mWorkSheet.UsedRange.Columns.Count
	lastRow = mWorkSheet.UsedRange.Rows.Count
	print("Laster Row : " + str(lastRow) + " colum num : " + str(lastCol))
	#mWorkSheet.Range("C2:C" + str(lastRow)).Select()  # 기업코드

	for i in range (2,  lastRow) :
		url = 'https://finance.naver.com/item/sise.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)
		try :
			table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함
			table_df = table_df_list[1]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
			df = pd.DataFrame(table_df)
			# df.iloc = row, column, 0부터 시작
			mWorkSheet.Cells(i, 10).Value = df.iloc[0][1] #현재가
			mWorkSheet.Cells(i, 11).Value = str(df.iloc[1][1]) + " " + str(df.iloc[2][1]) #전일대비 + 등락률(%)
		except :
			print('없는 기업 확인 필요= ' + str(mWorkSheet.Cells(i,4).Value))
			continue

		print(str(i) + '/'+ str(lastRow) + ' 기업명 = ' + str(mWorkSheet.Cells(i,4).Value))
		#print(df.head()) #print dataframe data
		#print(df.shape) #get row, column count

		url = 'https://finance.naver.com/item/main.nhn?code=' + str(mWorkSheet.Cells(i,3).Value)

		table_df_list = pd.read_html(url, encoding='euc-kr')
		table_df = table_df_list[5] #주요 재무제표
		df = pd.DataFrame(table_df)

		mWorkSheet.Cells(i, 12).Value = df.iloc[2][1] #상장주식수
		mWorkSheet.Cells(i, 13).Value = "=J"+ str(i) + "*L" + str(i) + "/100000000"   #시가총액(억) = 상장주식수 * 현재가

		table_df = table_df_list[7]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 141).Value = df.iloc[1][0] #52주 최고
		mWorkSheet.Cells(i, 142).Value = df.iloc[1][1] #52주 최저

		if mWorkSheet.Cells(i, 9).Value == str('SPAC') or mWorkSheet.Cells(i, 9).Value == str('리츠') :
			continue

		table_df = table_df_list[3] #주요 재무제표
		#table_df.columns = table_df.columns.droplevel(2)
		df = pd.DataFrame(table_df)
		df.fillna(0)

		if pd.isna(df.iloc[10,4]) or (str(df.iloc[10,4]) == str('-')) :
			#PER 어닝 이 없는 경우 직접 PER을 구한다.
			vTempTotalEPS = 0
			vTmpQtrCnt = 0
			if pd.isna(df.iloc[9,6]) or (str(df.iloc[9,6]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][6])
			if pd.isna(df.iloc[9,7]) or (str(df.iloc[9,7]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][7])
			if pd.isna(df.iloc[9,8]) or (str(df.iloc[9,8]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][8])
			if pd.isna(df.iloc[9,9]) or (str(df.iloc[9,9]) == str('-')) :
				vTempTotalEPS += 0
			else :
				vTmpQtrCnt += 1
				vTempTotalEPS += float(df.iloc[9][9])
			#print("vTempTotalEPS = " + str(vTempTotalEPS))
			if vTempTotalEPS != 0 :
				vTempTotalEPS = float(mWorkSheet.Cells(i, 10).Value / float(vTempTotalEPS))
            	# PER = 주가 / 주당순이익(EPS)
			mWorkSheet.Cells(i, 15).Value = vTempTotalEPS
			#print("estimated EPS : " + str(vTempTotalEPS))
		else :
			mWorkSheet.Cells(i, 15).Value =  df.iloc[10][4]

		if pd.isna(df.iloc[9,9]) or (str(df.iloc[9,9]) == str('-')) :
			vTempPER = 0
		else :
			vTempEPS = int(df.iloc[9][9])   #  최근 마지막 실적 2020.09 EPS
			vTempEPS = vTempEPS*4
			if vTempEPS != 0 :				
				vTempPER = float(mWorkSheet.Cells(i, 10).Value / float(vTempEPS))
			else :
				vTempPER = 0
		mWorkSheet.Cells(i, 16).Value = vTempPER  # 최근 분기 대비 PER
		#print("estimated quarter PER : " + str(vTempPER))

		if pd.isna(df.iloc[12,4])  or (str(df.iloc[12,4]) == str('-')) :
			# PBR 어닝이 없을 경우 직전 4분기 평균을 사용한다.
			vTmpQtrCnt = 0
			vTempPBR = 0
			if pd.isna(df.iloc[12,6]) or (str(df.iloc[12,6]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][6])
			if pd.isna(df.iloc[12,7]) or (str(df.iloc[12,7]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][7])
			if pd.isna(df.iloc[12,8]) or (str(df.iloc[12,8]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][8])
			if pd.isna(df.iloc[12,9]) or (str(df.iloc[12,9]) == str('-')) :
				vTempPBR += 0
			else :
				vTmpQtrCnt += 1
				vTempPBR += float(df.iloc[12][9])
			#print("estimated vTempPBR : " + str(vTempPBR))
			if vTmpQtrCnt != 0 :
				vTempPBR = float(vTempPBR/float(vTmpQtrCnt))
			mWorkSheet.Cells(i, 17).Value = vTempPBR
		else :
			mWorkSheet.Cells(i, 17).Value = df.iloc[12][4] # PBR
		#print("estimated PBR : " + str(mWorkSheet.Cells(i, 17).Value))

		if pd.isna(df.iloc[5,4]) :
			# ROE 어닝이 없을 경우 직전 4분기 평균을 사용한다.
			vTmpQtrCnt = 0
			vTempROE = 0
			if pd.isna(df.iloc[5,6]) or (str(df.iloc[5,6]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][6])
			if pd.isna(df.iloc[5,7]) or (str(df.iloc[5,7]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][7])
			if pd.isna(df.iloc[5,8]) or (str(df.iloc[5,8]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][8])
			if pd.isna(df.iloc[5,9]) or (str(df.iloc[5,8]) == str('-')) :
				vTempROE += 0
			else :
				vTmpQtrCnt += 1
				vTempROE += float(df.iloc[5][9])
			#print("estimated vTempROE : " + str(vTempROE))
			if vTmpQtrCnt != 0 :
				vTempROE = float(vTempROE/float(vTmpQtrCnt))
			mWorkSheet.Cells(i, 18).Value = vTempROE
		else :
			mWorkSheet.Cells(i, 18).Value = df.iloc[5][4] # ROE
			#print("estimated ROE : " + str(mWorkSheet.Cells(i, 18).Value))

		table_df = table_df_list[4]
		df = pd.DataFrame(table_df)
		mWorkSheet.Cells(i, 14).Value = df.iloc[4][1] #외국인비율(%)

		if i % 100 == 0 :
			mWorkBook.Save() #work around for excel fault

def run_Test_code() :
	#url = 'https://finance.naver.com/item/main.nhn?code=326030'
	#url = 'https://finance.naver.com/item/main.nhn?code=006360'
	url = 'https://finance.naver.com/item/coinfo.nhn?code=122450'
	try :
		table_df_list = pd.read_html(url, encoding='euc-kr')
		#table_df.columns = table_df.columns.droplevel(2)
		table_df = table_df_list[0]  # 리스트 중에서 원하는 데이터프레임 한개를 가져온다
		print('table_df_list[0]')
		df = pd.DataFrame(table_df)
		print(table_df)
	except :
		print('bad URL ')

def run_Test_javascript_code() :
	#pip install selenium
	#from selenium import webdriver
	browser = webdriver.Chrome()
	browser.get("https://www.python.org/")
	nav = browser.find_element_by_id("mainnav")
	print(nav.text)


filename = "export_Data.xlsx"
filepath = os.path.abspath(filename)
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

mWorkBook = excel.Workbooks.Open(filepath)
mWorkSheet = mWorkBook.Worksheets('RawData')
mWorkSheet.Select()

#get_company_list_full()
get_company_list_value()
#def run_each_company_data(company_code) :

#test code start
#run_Test_code()
#test code end
mWorkBook.Save()
mWorkBook.Close()
