#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os

hearthstonePath = "D:\misc\Hearthstone" # end without \

# extract from hearthstone
def extractFromHearthstone():
  cmd = ".\disunity_v0.3.4\disunity.bat extract %s\Data\Win\cardxml0.unity3d" % (hearthstonePath)
  os.system(cmd)


if __name__ == '__main__':
  extractFromHearthstone()





