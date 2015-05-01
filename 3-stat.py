#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook

wb = load_workbook(filename='cards-result.xlsx')
ws = wb['Sheet1']

kv = {u'傳說': 'lv1600', u'史詩': 'lv400', u'精良': 'lv100', u'普通': 'lv40', u'基本': 'lv0', None: 'lv0'}
stat = {}
for row in ws.rows:
  source = row[14].value
  level = row[18].value
  if source in [u'經典', u'GVG']:
    if not stat.has_key(source):
      stat[source] = {'total': 0, 'total_dust': 0, 'need': 0, 'need_dust': 0, 'lv0': 0, 'lv40': 0, 'lv100': 0, 'lv400': 0, 'lv1600': 0, 'tlv0': 0, 'tlv40': 0, 'tlv100': 0, 'tlv400': 0, 'tlv1600': 0}
    stat[source]['total'] += row[1].value
    stat[source]['total_dust'] += row[2].value
    stat[source]['need'] += row[5].value
    stat[source]['need_dust'] += row[6].value
    stat[source][kv[level]] += row[5].value
    stat[source]["t" + kv[level]] += row[1].value

for it in [u'經典', u'GVG']:
  print '===================================='
  foo = stat[it]
  print u'%s卡 缺: %s/%s(%s%%) 張 %s/%s(%s%%) 塵' % (it, foo['need'], foo['total'], foo['need'] * 100 / foo['total'], foo['need_dust'], foo['total_dust'], foo['need_dust'] * 100 / foo['total_dust'])
  print u'最多 %s 包' % (foo['need_dust'] / 40)
  print u'傳說 %s/%s 張 - 史詩 %s/%s 張 - 精良 %s/%s 張 - 普通 %s/%s 張' % (foo['lv1600'], foo['tlv1600'], foo['lv400'], foo['tlv400'], foo['lv100'], foo['tlv100'], foo['lv40'], foo['tlv40'])
  


