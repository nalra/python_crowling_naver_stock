#!/usr/bin/env python

import feedparser  # pip install feedparser
from pprint import pprint
from dateparser import parser # pip install python-dateutil


url = "http://dart.fss.or.kr/api/todayRSS.xml"

rss_info  = feedparser.parse(url)
disclosure_list = rss_info['entries']

for disclosure in disclosure_list:
	company = disclosure['authors']  # author: '티맵모빌리티'
	link = disclosure['link'] # can be a unique id role.
	published = parser.parse(disclosure['published'])
	year = published.year
	month = published.month
	day = published.day
	hour = published.hour + 9
	minute = published.minute
	date_info = f'{year}년 {month}월 {day}일 {hour}시 {minutes}분'
	title = disclosure['title']
	print(company, title, date_info)