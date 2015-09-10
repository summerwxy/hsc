#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from colorama import init
import random

PACK_NAME = [u'經典', u'GVG', u'TGT']
stat = {}

def readExcel():
  wb = load_workbook(filename='cards-result.xlsx')
  ws = wb['Sheet1']

  kv = {u'傳說': 'legendary', u'史詩': 'epic', u'精良': 'rare', u'普通': 'common', u'基本': 'basic', None: 'basic'}
  for row in ws.rows:
    source = row[14].value
    level = row[16].value
    if source in PACK_NAME and row[1].value:
      if not source in stat:
        stat[source] = {'total': 0, 'total_dust': 0, 'need': 0, 'need_dust': 0, 'basic': 0, 'common': 0, 'rare': 0, 'epic': 0, 'legendary': 0, 'tbasic': 0, 'tcommon': 0, 'trare': 0, 'tepic': 0, 'tlegendary': 0}
      stat[source]['total'] += row[1].value
      stat[source]['total_dust'] += row[2].value
      stat[source]['need'] += row[5].value
      stat[source]['need_dust'] += row[6].value
      stat[source][kv[level]] += row[5].value
      stat[source]["t" + kv[level]] += row[1].value


def printStat():
  alldust = 0
  for it in PACK_NAME:
    print('\033[1;31m====================================\033[m')
    foo = stat[it]
    alldust = alldust + foo['need_dust']
    print('%s: %s/%s(%s%%) Cards | %s/%s(%s%%) Dusts' % (it, foo['need'], foo['total'], round(foo['need'] * 100 / foo['total'], 2), foo['need_dust'], foo['total_dust'], round(foo['need_dust'] * 100 / foo['total_dust'], 2)))
    print('\033[0;33mLegendary %s/%s\033[m | \033[1;35mEpic %s/%s\033[m | \033[1;36mRare %s/%s\033[m | \033[1;37mCommon %s/%s\033[m' % (foo['legendary'], foo['tlegendary'], foo['epic'], foo['tepic'], foo['rare'], foo['trare'], foo['common'], foo['tcommon']))

  havedust = 8300 + 1570
  print('\033[1;31m-------------------------------------------\033[m')
  print("DUST ALL: %s / HAVE: %s / NEED: %s" % (alldust, havedust, alldust - havedust))
  print('\033[1;31m-------------------------------------------\033[m')
  
  stat['alldust'] = alldust
  stat['havedust'] = havedust

# http://www.carddust.com/
chanceLegendary = 0.0119;
chanceEpic = 0.0476;
chanceRare = 0.238;
chanceCommon = 0.70254;


def openPack(pack):
  times = 0
  commonMark = True
  while(times < 5):
    # rarity
    r = random.random()
    rarity = (r < chanceLegendary) and 'legendary' or (r < chanceLegendary + chanceEpic) and 'epic' or (r < chanceLegendary + chanceEpic + chanceRare) and 'rare' or 'common'
    if rarity != 'common':
      commonMark = False
    if commonMark and times == 4: 
      continue 
    # gold
    r = random.random()
    gold = False
    if rarity in ['legendary', 'epic', 'rare'] and r < 0.05:
      gold = True
    if rarity in ['common'] and r < 0.02:
      gold = True
    # get
    r = random.random()
    foo = stat[pack]
    get = False
    if r < foo[rarity] / foo['t' + rarity]:
      get = True
    # calc dust
    if get:
      foo[rarity] = foo[rarity] - 1
      foo['need'] = foo['need'] - 1
      this_dust = {'legendary': 1600, 'epic': 400, 'rare': 100, 'common': 40}[rarity]
      foo['need_dust'] = foo['need_dust'] - this_dust
    else:
      this_dust = {'legendaryTrue': 1600, 'epicTrue': 400, 'rareTrue': 100, 'commonTrue': 50, 'legendaryFalse': 400, 'epicFalse': 100, 'rareFalse': 20, 'commonFalse': 5}[rarity + str(gold)]
      stat['havedust'] = stat['havedust'] + this_dust
    times += 1

def getAllPack(pack):
  print(pack)
  times = 0
  while stat['havedust'] < stat[pack]['need_dust']:
    openPack(pack)
    times += 1
  print(times)
  print(stat)

# TODO: how to open pack is better


if __name__ == '__main__':
  init() # use color console
  readExcel()
  printStat()
  getAllPack(PACK_NAME[2])


