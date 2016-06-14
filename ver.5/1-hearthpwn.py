#!/usr/bin/env python
# -*- coding: utf-8 -*-

import urllib.request
from lxml import etree, cssselect
import shutil
from colorama import init

def grabit():
  url = "http://www.hearthpwn.com/members/lin0_o/collection"
  res = urllib.request.urlopen(url)
  html = res.read()
  return html

def grabit2file():
  url = "http://www.hearthpwn.com/members/lin0_o/collection"
  res = urllib.request.urlopen(url)
  f = open(r'Hearthpwn.html', 'wb')
  shutil.copyfileobj(res, f)

def loadfile():
  f = open(r'Hearthpwn.html', 'r')
  html = f.read()
  return html

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

def parseit(html):
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



if __name__ == '__main__':
  if input("Download from Hearthpwn? (y/N)") == 'y':
    grabit2file()
    print("Download DONE!")
  html = loadfile()
  #html = grabit()
  data = parseit(html)
  printit(data)

  cards = parsecard(html)
  print(len(cards))



