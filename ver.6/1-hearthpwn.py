#!/usr/bin/env python
# -*- coding: utf-8 -*-

import urllib.request
from lxml import etree, cssselect
import shutil
from colorama import init
import requests
import codecs
import json
import xlsxwriter
import re

def grabit2file():
  url = "https://api.hearthstonejson.com/v1/latest/zhTW/cards.collectible.json"
  headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
  result = requests.get(url, headers=headers)
  f = open(r'zhTW.cards.collectible.json', 'wb')
  f.write(result.content)
  f.close()


def loadfile():
  f = codecs.open(r'zhTW.cards.collectible.json', 'r', 'utf-8')
  return f.read()


def extractInfo(div):
  id = div.get('id')[20:]
  dust_span = div.cssselect('div.collection-dust-required span.type-switch')[0]
  need_dust = dust_span.get('data-combined')
  dust_span = div.cssselect('ul.rarity-stats li span.type-switch')
  fqty = dust_span[0].get('data-combined') # free
  cqty = dust_span[1].get('data-combined') # common
  rqty = dust_span[2].get('data-combined') # rare
  eqty = dust_span[3].get('data-combined') # epic
  lqty = dust_span[4].get('data-combined') # legendary
  return id, need_dust, fqty, cqty, rqty, eqty, lqty

def parseit(jsonobj):
  workbook = xlsxwriter.Workbook('cards-result.xlsx')

  header = workbook.add_format({'bg_color': '#000000', 'font_color': '#ffffff', 'bold': True, 'border': 1})
  black = workbook.add_format({'bg_color': '#5a5a5a', 'border': 1})
  basic = workbook.add_format({'bg_color': '#d7e4bc', 'border': 1})
  green = workbook.add_format({'bg_color': '#c2d69a', 'border': 1})
  orange = workbook.add_format({'bg_color': '#fcd5b4', 'border': 1})
  blue = workbook.add_format({'bg_color': '#b6dde8', 'border': 1})
  red = workbook.add_format({'bg_color': '#e6b9b8', 'border': 1})
  white = workbook.add_format({'bg_color': '#ffffff', 'border': 1})

  worksheet = workbook.add_worksheet()

  title = ['id', 'playerClass', 'name', 'text', 'rarity', 'type', 'cost', 'attack', 'health', 'collectible', 'set', 'faction', 'artist', 'flavor', 'mechanics', 'dust', 'dbfId']
  for i in range(len(title)):
    worksheet.write(0, i, title[i])  
  
  for i in range(len(jsonobj)):
    card = jsonobj[i]
    worksheet.write(i + 1, 0, card['id']) 
    worksheet.write(i + 1, 1, card['playerClass']) 
    worksheet.write(i + 1, 2, card['name']) 
    worksheet.write(i + 1, 3, ('text' in card) and remove_tags(card['text']) or '') 
    worksheet.write(i + 1, 4, card['rarity']) 
    worksheet.write(i + 1, 5, card['type']) 
    worksheet.write(i + 1, 6, ('cost' in card) and card['cost'] or '') 
    worksheet.write(i + 1, 7, ('attack' in card) and card['attack'] or '') 
    worksheet.write(i + 1, 8, ('health' in card) and card['health'] or '') 
    worksheet.write(i + 1, 9, card['collectible']) 
    worksheet.write(i + 1, 10, card['set']) 
    worksheet.write(i + 1, 11, ('faction' in card) and card['faction'] or '') 
    worksheet.write(i + 1, 12, ('artist' in card) and card['artist'] or '') 
    worksheet.write(i + 1, 13, ('flavor' in card) and card['flavor'] or '') 
    worksheet.write(i + 1, 14, ('mechanics' in card) and ", ".join(card['mechanics']) or '') 
    worksheet.write(i + 1, 15, ('dust' in card) and ", ".join([str(x) for x in card['dust']]) or '') 
    worksheet.write(i + 1, 16, card['dbfId']) 

  #ws = [5, 3, 5, 3, 5, 3, 5, 3, 5, 10, 22, 22, 3, 3, 3, 5, 5, 5.5, 5, 5, 60]
  #for i in range(len(ws)):
  #  worksheet.set_column(i, i, ws[i])

  #worksheet.autofilter('A1:U1')
  workbook.close()

  """
  title = ['dust', 'total', 'total', 'have', 'have', 'need', 'need', 'wish', 'wish', 'id', 'ename', 'name', 'mp', 'atk', 'hp', 'set', 'type', 'rarity', 'class', 'race', 'effect']
  for i in range(len(title)):
    worksheet.write(0, i, title[i])

  i = 1
  for k in tdict:
    card = tdict[k]
    cid = card['CardId']
    cset = card['CardSet']
    ctype = card['Type']
    crarity = card['Rarity']
    cclass = card['Class']
    crace = card['Race']
  """


  """
  result = []
  parser = etree.HTMLParser()
  html = etree.fromstring(html, parser)
  select = cssselect.CSSSelector(".set-container")
  divs = select(html)
  if not divs: # no data page
    return
  for i in range(len(divs)):
    id, dust, f, c, r, e, l = extractInfo(divs[i]) # id, dust, free, common, rare, epic, legendary
    result.append({'id': id, 'dust': dust, 'free': f, 'common': c, 'rare': r, 'epic': e, 'legendary': l})
  return result
  """

def printit(data):
  init() # use color console

  for i in range(len(data)):
    setc = data[i]
    if setc['id'] in ['Classic', 'GvGx', 'TGT', 'WOG']:
      print("")
      print('\033[1;31m== %s ==================================\033[m' % setc['id'])
      l = calcq(setc['legendary'])
      e = calcq(setc['epic'])
      r = calcq(setc['rare'])
      c = calcq(setc['common'])
      cs, ds = sumq(setc)
      print('\033[0;37m%s Cards\033[m | \033[0;37m%s Dusts\033[m' % (cs, ds))
      print('\033[0;33mL %s\033[m | \033[1;35mE %s\033[m | \033[1;36mR %s\033[m | \033[1;37mC %s\033[m' % (l, e, r, c))




def calcq(s):
  foo = [int(x) for x in s.split('/')]
  return '%s(%s%%)' % (s, round(foo[0] * 100 / foo[1], 2))


def sumq(ss):
  l = [int(x) for x in ss['legendary'].split('/')]
  e = [int(x) for x in ss['epic'].split('/')]
  r = [int(x) for x in ss['rare'].split('/')]
  c = [int(x) for x in ss['common'].split('/')]
  fc = l[0] + e[0] + r[0] + c[0]
  nc = l[1] + e[1] + r[1] + c[1]
  pc = round(fc * 100 / nc, 2)
  sc = '%s/%s(%s%%)' % (fc, nc, pc)
  fd = l[0] * 1600 + e[0] * 400 + r[0] * 100 + c[0] * 40
  nd = l[1] * 1600 + e[1] * 400 + r[1] * 100 + c[1] * 40
  pd = round(fd * 100 / nd, 2)
  sd = '%s/%s(%s%%)' % (fd, nd, pd)
  return sc, sd

def parsecard(html):
  cards = {}
  parser = etree.HTMLParser()
  html = etree.fromstring(html, parser)
  select = cssselect.CSSSelector('.card-image-item')
  divs = select(html) 
  if not divs:
    return
  for i in range(len(divs)):
    div = divs[i]
    name = div.get('data-card-name')
    card = {}
    if name in cards:
      card = cards[name]
    else:
      card['id'] = div.get('data-id')
      card['rarity'] = div.get('data-rarity')
      card['name'] = name
      card['mana'] = div.get('data-card-mana-cost')
      card['description'] = div.get('data-card-description')
      card['race'] = div.get('data-card-race')
      card['class'] = div.get('data-card-class')
      card['hp'] = div.get('data-card-hp')
      card['attack'] = div.get('data-card-attack')
      card['type'] = div.get('data-card-type')
      card['mechanics'] = div.get('data-card-mechanics')
      card['total'] = 0
      card['have'] = 0
    span = div.find('a').cssselect('span.inline-card-count')[0]
    foo = span.text.split('/')
    card['total'] = int(foo[1])
    card['have'] += int(foo[0])
    card['have'] = (card['total'] > card['have'] and [card['have']] or [card['total']])[0]
    gold = div.get('data-is-gold')
    if gold == 'True':
      card['gold_count'] = span.text
    else:
      card['normal_count'] = span.text
    cards[name] = card
  return cards

TAG_RE = re.compile(r'<[^>]+>')

def remove_tags(text):
  return TAG_RE.sub('', text)


if __name__ == '__main__':
  if input("Download from hearthstonejson? (y/N)") == 'y':
    grabit2file()
    print("Download DONE!")
  jsonstr = loadfile()
  jsonobj = json.loads(jsonstr)
  data = parseit(jsonobj)
  

  """
  data = parseit(html)
  printit(data)

  cards = parsecard(html)
  print(len(cards))
  """


