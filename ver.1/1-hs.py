#!/usr/bin/env python
# -*- coding: utf-8 -*-

import urllib
import urllib2
import os.path
import json
from lxml import etree, cssselect
import codecs


def parseCard(html_code, i):
  result = {}
  parser = etree.HTMLParser()
  html = etree.fromstring(html_code, parser)
  # card id
  result['id'] = i
  # chinese name
  select = cssselect.CSSSelector("table.table_out a span")
  if not select(html): # no data page
    return
  result['cname'] = select(html)[0].text.strip()
  # english name
  select = cssselect.CSSSelector("table.table_out a br")
  result['ename'] = select(html)[0].tail.strip()
  # image url
  select = cssselect.CSSSelector("#card_book_container img")
  result['img'] = select(html)[0].get('src')
  # others
  select = cssselect.CSSSelector("div table div table.table_out div table tr td")
  foo = select(html)


  result['class'] = foo[7].text
  #result['source'] = foo[9].text
  result['level'] = foo[8].text
  result['type'] = foo[9].text
  result['race'] = foo[10].text
  result['mp'] = foo[11].getchildren()[0].text
  result['atk'] = foo[12].getchildren()[0].text
  result['hp'] = foo[13].getchildren()[0].text
  bar = foo[15].find('div')
  if bar is not None:
    result['eeffect'] = bar.text
    bar.clear()
  else:
    result['eeffect'] = ''
  result['ceffect'] = "".join([t.strip() for t in foo[15].itertext() if t.strip()])
  result['desc'] = foo[17].text
  result['misc'] = "".join([t.strip() for t in foo[len(foo) - 1].itertext() if t.strip()])
  return result

def downloadCard(path, id):
  # TODO: check, download if not exist
  if not os.path.isfile('images\%s.jpg' % id):
    try:
      url = 'http://gametsg.techbang.com/hs/%s'
      image = urllib.URLopener()
      image.retrieve(url % path, "images\%s.jpg" % id)
      print '[download ok]'
    except:
      print '[download ERROR]'
  else:
    print '[exist]'

def runit():
  cards = []
  url = 'http://gametsg.techbang.com/hs/index.php?view=item&item=%s#detail'
  #ids = range(10001, 10529 + 1)         # card no. 10001 ~ 10529 基本 經典 獎勵 促銷
  #ids.extend(range(11001, 11085 + 1))   # card no. 11001 ~ 11085 NAXX
  #ids.extend(range(12001, 12144 + 1))   # card no. 12001 ~ 12144 GVG
  #ids.extend(range(12001, 12144 + 1))   # card no. 13001 ~ 13204 BRM
  source = 'BRM'
  ids = range(13001, 13204 + 1)

  for i in ids:
    # get page
    html = urllib2.urlopen(url % i).read()
    # parse data
    data = parseCard(html, i)
    if data:
      print i, data['cname'], 
      data['source'] = source
      # download card image
      downloadCard(data['img'], i)
      # dump to json
      cards.append(data)
    else:
      print i, 'PASS!'

    # for test
    if data and False:
      for it in data.items():
        print it[0], ' => ', it[1]
  # write to file
  f = codecs.open("cards.js", "w", "utf-8")
  foo = json.dumps(cards, ensure_ascii=False, encoding='utf8')
  f.write(foo)
  f.close()

if __name__ == '__main__':
  runit()

