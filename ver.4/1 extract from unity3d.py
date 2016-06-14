#!/usr/bin/env python
import re
from lxml.html import fromstring


def extract():
  cardxml = open('D:\\misc\\Hearthstone\\Data\Win\\cardxml0.unity3d', 'r', encoding='utf-8', errors='ignore').read()
  cards = re.findall(r'\<Entity.*?Entity\>', cardxml, re.DOTALL)
  i = 1
  setid = 1
  first_cardid = ''
  with open('cards.xml', 'w', encoding='utf-8') as f:
    for card in cards:
      # TODO: parse to another format
      doc = fromstring(card)
      cardid = doc.attrib['cardid']
      if i == 1:
        first_cardid = cardid
      elif cardid == first_cardid:
        setid += 1

      # 7 = tw, 11 / 12 = en, 15 = cn
      if setid == 7 and cardid == first_cardid:
        print(card)
        print(dir(doc))
        print([child.text for child in doc.iterchildren()])
        





      """
      if card.count("CS1h_001") > 0 and not i in [3873, 7745, 9681, 13553, 15489, 19361, 21297, 23233, 25169]:
        print(i)
        print(card)
      """

      """
      if i in range(1937, 3873):
        print(card)

      """
      i = i + 1



if __name__ == '__main__':
  extract()


