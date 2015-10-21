#!/usr/bin/env python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from colorama import init
from random import random
from copy import deepcopy

PACK_NAME = {'cls': u'經典', 'gvg': u'GVG', 'tgt': u'TGT'}

# http://www.carddust.com/
chanceLegendary = 0.0119;
chanceEpic = 0.0476;
chanceRare = 0.238;
chanceCommon = 0.70254;



def readExcel(havedust):
  data = {}
  wb = load_workbook(filename='cards-result.xlsx')
  ws = wb['Sheet1']

  kv = {u'傳說': 'legendary', u'史詩': 'epic', u'精良': 'rare', u'普通': 'common', u'基本': 'basic', None: 'basic'}
  for row in ws.rows:
    source = row[14].value
    level = row[16].value
    if source in sorted(PACK_NAME.values()) and row[1].value:
      key = [k for k, v in PACK_NAME.items() if v == source][0]
      if not source in data:
        data[source] = {'key': key, 'total': 0, 'total_dust': 0, 'need': 0, 'need_dust': 0, 'basic': 0, 'common': 0, 'rare': 0, 'epic': 0, 'legendary': 0, 'tbasic': 0, 'tcommon': 0, 'trare': 0, 'tepic': 0, 'tlegendary': 0}
      data[source]['total'] += row[1].value
      data[source]['total_dust'] += row[2].value
      data[source]['need'] += row[5].value
      data[source]['need_dust'] += row[6].value
      data[source][kv[level]] += row[5].value
      data[source]["t" + kv[level]] += row[1].value
  data['havedust'] = havedust
  return data


def printData(data):
  alldust = 0
  for key in sorted(PACK_NAME.keys()):
    it = PACK_NAME[key]
    print('\033[1;31m====================================\033[m')
    foo = data[it]
    alldust = alldust + foo['need_dust']
    print('%s: %s/%s(%s%%) Cards | %s/%s(%s%%) Dusts' % (it, foo['need'], foo['total'], round(foo['need'] * 100 / foo['total'], 2), foo['need_dust'], foo['total_dust'], round(foo['need_dust'] * 100 / foo['total_dust'], 2)))
    print('\033[0;33mLegendary %s/%s\033[m | \033[1;35mEpic %s/%s(%s)\033[m | \033[1;36mRare %s/%s(%s)\033[m | \033[1;37mCommon %s/%s(%s)\033[m' % (foo['legendary'], foo['tlegendary'], foo['epic'], foo['tepic'], round(foo['epic'] * 100 / foo['tepic'], 2), foo['rare'], foo['trare'], round(foo['rare'] * 100 / foo['trare'], 2), foo['common'], foo['tcommon'], round(foo['common'] * 100 / foo['tcommon'], 2)))

  print('\033[1;31m-------------------------------------------\033[m')
  print("DUST ALL: %s / HAVE: %s / NEED: %s" % (alldust, data['havedust'], alldust - data['havedust']))
  print('\033[1;31m-------------------------------------------\033[m')



def openOnePack(simData, key):
  pack = PACK_NAME[key]
  times = 0
  commonMark = True
  while(times < 5):
    # rarity
    r = random()
    rarity = (r < chanceLegendary) and 'legendary' or (r < chanceLegendary + chanceEpic) and 'epic' or (r < chanceLegendary + chanceEpic + chanceRare) and 'rare' or 'common'
    if rarity != 'common':
      commonMark = False
    if commonMark and times == 4: 
      continue 
    # gold
    r = random()
    gold = False
    if rarity in ['legendary', 'epic', 'rare'] and r < 0.05:
      gold = True
    if rarity in ['common'] and r < 0.02:
      gold = True
    # get
    r = random()
    foo = simData[pack]
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
      simData['havedust'] = simData['havedust'] + this_dust
    times += 1
  return simData

# open some pack until get all
def simOpenPack(data, key, times):
  pack = PACK_NAME[key]
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    while simData['havedust'] < simData[pack]['need_dust']:
      simData = openOnePack(simData, pack)
      opens += 1
    avg.append(opens)
  
  a = sum(avg) / len(avg)
  print(pack + ' -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))

def simOpenPackWithoutDust(data, key, times):
  pack = PACK_NAME[key]
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    while simData[pack]['need_dust'] > 0:
      simData = openOnePack(simData, pack)
      opens += 1
    avg.append(opens)
  
  a = sum(avg) / len(avg)
  print(pack + ' -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))
  

# open pack until get all card: classic -> gvg -> tgt
def simOpenAll_1(data, times):
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    while simData['havedust'] < sum([it['need_dust'] for it in simData.values() if type(it) == dict]):
      if simData[PACK_NAME['cls']]['need_dust'] > 0:
        simData = openOnePack(simData, 'cls')
      elif simData[PACK_NAME['gvg']]['need_dust'] > 0:
        simData = openOnePack(simData, 'gvg')
      elif simData[PACK_NAME['tgt']]['need_dust'] > 0:
        simData = openOnePack(simData, 'tgt')
      else:
        print("WTF!!!!")
      opens += 1
    avg.append(opens)
  a = sum(avg) / len(avg)
  print('OPEN ALL 1 -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))


# open pack until get all card: tgt -> gvg -> classic
def simOpenAll_2(data, times):
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    while simData['havedust'] < sum([it['need_dust'] for it in simData.values() if type(it) == dict]):
      if simData[PACK_NAME['tgt']]['need_dust'] > 0:
        simData = openOnePack(simData, 'tgt')
      elif simData[PACK_NAME['gvg']]['need_dust'] > 0:
        simData = openOnePack(simData, 'gvg')
      elif simData[PACK_NAME['cls']]['need_dust'] > 0:
        simData = openOnePack(simData, 'cls')
      else:
        print("WTF!!!!")
      opens += 1
    avg.append(opens)
  a = sum(avg) / len(avg)
  print('OPEN ALL 2 -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))

# open pack until get all card: common -> rare -> epic -> legendary by card
def simOpenAll_3(data, times):
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    while simData['havedust'] < sum([it['need_dust'] for it in simData.values() if type(it) == dict]):
      if sum([it['common'] for it in simData.values() if type(it) == dict]):
        foo = sorted([it for it in simData.values() if type(it) == dict], key=lambda d: d['common'])[-1]
        simData = openOnePack(simData, foo['key'])
      elif sum([it['rare'] for it in simData.values() if type(it) == dict]):
        foo = sorted([it for it in simData.values() if type(it) == dict], key=lambda d: d['rare'])[-1]
        simData = openOnePack(simData, foo['key'])
      elif sum([it['epic'] for it in simData.values() if type(it) == dict]):
        foo = sorted([it for it in simData.values() if type(it) == dict], key=lambda d: d['epic'])[-1]
        simData = openOnePack(simData, foo['key'])
      elif sum([it['legendary'] for it in simData.values() if type(it) == dict]):
        foo = sorted([it for it in simData.values() if type(it) == dict], key=lambda d: d['legendary'])[-1]
        simData = openOnePack(simData, foo['key'])
      else:
        print("WTF!!" + str(opens))
      opens += 1
    avg.append(opens)
  a = sum(avg) / len(avg)
  print('OPEN ALL 3 -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))

# open pack until get all card: common -> rare -> epic -> legendary by rate
def simOpenAll_4(data, times):
  avg = []
  for i in range(times):
    simData = deepcopy(data)
    opens = 0
    sortBy = ''
    while simData['havedust'] < sum([it['need_dust'] for it in simData.values() if type(it) == dict]):
      if sum([it['common'] for it in simData.values() if type(it) == dict]):
        sortBy = 'common'
      elif sum([it['rare'] for it in simData.values() if type(it) == dict]):
        sortBy = 'rare'
      elif sum([it['epic'] for it in simData.values() if type(it) == dict]):
        sortBy = 'epic'
      elif sum([it['legendary'] for it in simData.values() if type(it) == dict]):
        sortBy = 'legendary'
      else:
        print("WTF!!" + str(opens))
      foo = sorted([it for it in simData.values() if type(it) == dict], key=lambda d: d[sortBy] / d['t' + sortBy])[-1]
      simData = openOnePack(simData, foo['key'])
      opens += 1
    avg.append(opens)
  a = sum(avg) / len(avg)
  print('OPEN ALL 4 -> ' + str(a) + ' MIN: ' + str(min(avg)) + ' MAX: ' + str(max(avg)))

def ifOpenPack(data, key, times):

  simData = deepcopy(data)
  for i in range(times): 
    simData = openOnePack(simData, key)
  print(">>>>>>>>>>>>>> after %s packs <<<<<<<<<<<<<" % (times))
  printData(simData)


if __name__ == '__main__':
  init() # use color console
  # wxy dust
  havedust = 9445 + 1575
  # chicken dust
  #havedust = 45 + 20
  data = readExcel(havedust)

  # if open 60 pack
  #ifOpenPack(data, 'cls', 0)

  
  printData(data)
  t = 100
  """
  simOpenPack(data, 'cls', t)
  simOpenPack(data, 'gvg', t)
  simOpenPack(data, 'tgt', t)
  simOpenPackWithoutDust(data, 'cls', t)
  simOpenPackWithoutDust(data, 'gvg', t)
  simOpenPackWithoutDust(data, 'tgt', t)
  simOpenAll_1(data, t)
  simOpenAll_2(data, t)
  simOpenAll_3(data, t)
  """
  simOpenAll_4(data, t)


