#!/usr/bin/env python

import feedparser  # pip install feedparser
from pprint import pprint
from dateutil import parser # pip install python-dateutil
from time import sleep

from feedparser.exceptions import ThingsNobodyCaresAboutButMe # pip install dateparser
import telegram #pip install python-telegram-bot

disclosure_list_old = []
TOKEN = ''
CHAT_ID = ''
bot = telegram.Bot(token=TOKEN)

while True:

	url = "http://dart.fss.or.kr/api/todayRSS.xml"

	rss_info  = feedparser.parse(url)
	disclosure_list = rss_info['entries']

	if len(disclosure_list_old) == 0:
		for disclosure in disclosure_list:
			company = disclosure['author']  # author: 'ex: 삼성전자(주)'
			link = disclosure['link'] # can be a unique id role.
			published = parser.parse(disclosure['published'])
			year = published.year
			month = published.month
			day = published.day
			hour = published.hour + 9
			minute = published.minute
			date_info = f'{year}년 {month}월 {day}일 {hour}시 {minute}분'
			title = disclosure['title']
			message = f'{title}\n{link}\n{date_info}'

			if '공급계약체결' or '무상증자결정' or '공개매수신고서' or '소송등의제기ㆍ신청(경영권분쟁소송)' in title:  #'유무상증자' 는 제외해야 함.
				print(message)
				#print(company, title, date_info)
	else:
		for disclosure in disclosure_list:
			if disclosure['link'] == disclosure_list_old[0]['link']:
				break
			company = disclosure['author']  # author: 'ex: 삼성전자(주)'
			link = disclosure['link'] # can be a unique id role.
			published = parser.parse(disclosure['published'])
			year = published.year
			month = published.month
			day = published.day
			hour = published.hour + 9
			minute = published.minute
			date_info = f'{year}년 {month}월 {day}일 {hour}시 {minute}분'
			title = disclosure['title']
			if '공급계약체결' or '무상증자결정' or '공개매수신고서' or '소송등의제기ㆍ신청(경영권분쟁소송)' in title:  #'유무상증자' 는 제외해야 함.
				print(company, title, date_info)

	disclosure_list_old = disclosure_list
	sleep(10)

def print_disclosure() :