#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import codecs
import xlsxwriter

def runit():
  f = codecs.open("cs.js", "r", "utf-8")
  s = "".join(f.readlines())
  cards = json.loads(s)

  workbook = xlsxwriter.Workbook('hs.xlsx')
  worksheet = workbook.add_worksheet()

  cols = [u'ename', u'img', u'level', u'hp', u'ceffect', u'misc', u'atk', u'id', u'race', u'mp', u'cname', u'eeffect', u'type', u'class', u'desc']
  cols = [u'id', u'cname', u'mp', u'atk', u'hp', u'type', u'class', u'race', u'level', u'ceffect', u'desc', u'misc', u'ename', u'eeffect', u'img']
  j = 0
  for col in cols:
    worksheet.write(0, j, col)
    j = j + 1

  i = 1
  for card in cards:
    j = 0
    for col in cols:
      worksheet.write(i, j, card[col])
      j = j + 1
    i = i + 1
  workbook.close()


if __name__ == '__main__':
  runit()
  

