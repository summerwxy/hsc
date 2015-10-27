#!/usr/bin/env python
# -*- coding: utf-8 -*-
import urllib.request
import shutil

# 下載這裡的檔案
# https://github.com/Epix37/Hearthstone-Deck-Tracker/tree/master/Hearthstone%20Deck%20Tracker/Files

def downloadXml(lang):
  print('downloading {0}'.format(lang))
  url = "https://raw.githubusercontent.com/Epix37/Hearthstone-Deck-Tracker/master/Hearthstone%20Deck%20Tracker/Files/cardDB.{0}.xml".format(lang)
  with urllib.request.urlopen(url) as response, open(r'xml\{0}.xml'.format(lang), 'wb') as out_file:
    shutil.copyfileobj(response, out_file)




if __name__ == '__main__':
  downloadXml('enUS')
  downloadXml('zhTW')
  downloadXml('zhCN')





