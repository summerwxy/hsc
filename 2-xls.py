#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import codecs
import xlsxwriter

def runit():
  f = codecs.open("cards-classic.js", "r", "utf-8")
  s = "".join(f.readlines())
  cards = json.loads(s)
  f = codecs.open("cards-naxx.js", "r", "utf-8")
  s = "".join(f.readlines())
  cards = cards + json.loads(s)
  f = codecs.open("cards-gvg.js", "r", "utf-8")
  s = "".join(f.readlines())
  cards = cards + json.loads(s)
  f = codecs.open("cards-brm.js", "r", "utf-8")
  s = "".join(f.readlines())
  cards = cards + json.loads(s)

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

  cols = [u'id', u'cname', u'mp', u'atk', u'hp', u'source', u'type', u'class', u'race', u'level', u'ceffect', u'desc', u'misc', u'ename', u'eeffect', u'img']
  worksheet.set_row(0, 13.5, header)

  j = 0
  for title in [u'dust', u'total', u'total', u'have', u'have', u'need', u'need', u'wish', u'wish']:
    worksheet.write(0, j, title)
    j = j + 1
  for col in cols:
    worksheet.write(0, j, col)
    j = j + 1

  i = 1
  for card in cards:
    dust = (str(card[u'id']) in token and [0] or card[u'level'] == u'傳說' and [1600] or card[u'level'] == u'史詩' and [400] or card[u'level'] == u'精良' and [100] or card[u'level'] == u'普通' and [40] or [0])[0]
    worksheet.write(i, 0, dust) 
    mys = my.has_key(str(card[u'id'])) and my[str(card[u'id'])] or {'h': 0, 'w': 0}
    t = (str(card[u'id']) in token and [0] or mys.has_key(u't') and [mys[u't']] or card[u'level'] == u'傳說' and [1] or card[u'level'] == u'基本' and [0] or [2])[0]
    h = (card[u'source'] in chap or str(card[u'id']) in have) and t or mys['h']
    w = mys['w'] and mys['w'] - h or 0
    w = w > 0 and w or 0

    worksheet.write(i, 1, t)
    worksheet.write(i, 2, t * dust)
    worksheet.write(i, 3, h)
    worksheet.write(i, 4, h * dust)
    worksheet.write(i, 5, t - h)
    worksheet.write(i, 6, (t - h) * dust)
    worksheet.write(i, 7, w)
    worksheet.write(i, 8, w * dust)
    j = 9
    for col in cols:
      if col == u'desc':
        worksheet.write(i, j, card[col].strip())
      else:
        worksheet.write(i, j, card[col])
      j = j + 1
    # setting bg color
    format = white

    if str(card[u'id']) in token:
      format = black
    elif card[u'level'] == u'基本' or card[u'source'] in chap:
      format = basic
    elif h >= t:
      format = green
      if card[u'level'] == u'傳說':
        have.append(str(card[u'id']))
    elif card[u'level'] == u'傳說':
      format = orange
    elif t - h == 1:
      format = blue
    elif t - h == 2:
      format = red

    worksheet.set_row(i, 13.5, format)
    i = i + 1

  ws = [5, 3, 5, 3, 5, 3, 5, 3, 5, 5.5, 22, 3, 3, 3, 5, 5, 5.5, 5, 5, 40, 40]
  for i in range(len(ws)):
    worksheet.set_column(i, i, ws[i])


  workbook.close()

# get all chap
chap = ['NAXX', 'BRM']

# (t)otal (h)ave (w)ish
my = {
  # classic
  '10008': {'h': 1, 'w': 0}, # 奧多爾保安官
  '10009': {'h': 0, 'w': 1}, # 雅立史卓莎
  '10015': {'h': 0, 'w': 2}, # 知識古樹
  '10019': {'h': 1, 'w': 2}, # 遠古看守者
  '10031': {'h': 0, 'w': 1}, # 大法師安東尼達斯
  '10032': {'h': 1, 'w': 2}, # 銀白指揮官
  '10036': {'h': 0, 'w': 2}, # 護甲鍛造師
  '10042': {'h': 0, 'w': 2}, # 復仇之怒
  '10048': {'h': 0, 'w': 2}, # 末日災厄
  '10049': {'h': 0, 'w': 0}, # 迦頓男爵
  '10055': {'h': 0, 'w': 1}, # 狂野怒火
  '10067': {'h': 0, 'w': 2}, # 血騎士
  '10070': {'h': 0, 'w': 1}, # 血法師薩爾諾斯
  '10077': {'h': 1, 'w': 0}, # 鬥毆
  '10081': {'h': 0, 'w': 1}, # 綠皮船長
  '10082': {'h': 0, 'w': 0}, # 船長的鸚鵡
  '10084': {'h': 0, 'w': 1}, # 塞納留斯
  '10100': {'h': 0, 'w': 2}, # 法術反制
  '10128': {'h': 1, 'w': 2}, # 末日錘
  '10129': {'h': 1, 'w': 0}, # 末日預言者
  '10140': {'h': 0, 'w': 2}, # 鷹角弓
  '10141': {'h': 0, 'w': 2}, # 土元素
  '10144': {'h': 0, 'w': 1}, # 艾德溫·范克里夫
  '10151': {'h': 1, 'w': 2}, # 剔骨
  '10154': {'h': 1, 'w': 0}, # 爆裂射擊
  '10155': {'h': 1, 'w': 2}, # 爆炸陷阱
  '10157': {'h': 1, 'w': 2}, # 無面操縱者
  '10161': {'h': 0, 'w': 1}, # 視界術
  '10162': {'h': 1, 'w': 0}, # 惡魔守衛
  '10178': {'h': 1, 'w': 2}, # 照明彈
  '10180': {'h': 1, 'w': 2}, # 自然之力
  '10191': {'h': 0, 'w': 2}, # 加基森拍賣師
  '10193': {'h': 0, 'w': 0}, # 傑爾賓·梅卡托克
  '10203': {'h': 0, 'w': 0}, # 戈魯爾
  '10209': {'h': 0, 'w': 1}, # 哈里遜·瓊斯
  '10220': {'h': 0, 'w': 1}, # 霍格
  '10226': {'h': 0, 'w': 2}, # 神聖憤怒
  '10231': {'h': 1, 'w': 0}, # 飢餓的螃蟹
  '10237': {'h': 0, 'w': 1}, # 伊利丹·怒風
  '10243': {'h': 1, 'w': 2}, # 受傷的大劍師
  '10255': {'h': 0, 'w': 0}, # 綁匪
  '10257': {'h': 0, 'w': 1}, # 克洛許王
  '10258': {'h': 0, 'w': 1}, # 穆克拉
  '10266': {'h': 1, 'w': 0}, # 聖療術
  '10268': {'h': 0, 'w': 1}, # 炸雞勇者
  '10275': {'h': 0, 'w': 2}, # 閃電風暴
  '10282': {'h': 0, 'w': 1}, # 賈拉克瑟斯領主
  '10284': {'h': 0, 'w': 1}, # 博學行者阿洲
  '10289': {'h': 0, 'w': 1}, # 瑪里苟斯
  '10290': {'h': 1, 'w': 0}, # 魔法成癮者
  '10291': {'h': 0, 'w': 2}, # 法力之潮圖騰
  '10298': {'h': 0, 'w': 2}, # 群體驅魔
  '10301': {'h': 1, 'w': 0}, # 劍類鍛造大師
  '10305': {'h': 0, 'w': 0}, # 米歐浩斯·曼納斯頓
  '10308': {'h': 1, 'w': 2}, # 精神控制技師
  '10312': {'h': 0, 'w': 2}, # 心理遊戲
  '10316': {'h': 0, 'w': 2}, # 誤導
  '10319': {'h': 0, 'w': 2}, # 熔火巨人
  '10324': {'h': 0, 'w': 2}, # 山嶺巨人
  '10329': {'h': 0, 'w': 2}, # 魚人招潮者
  '10331': {'h': 0, 'w': 2}, # 魚人隊長
  '10333': {'h': 0, 'w': 0}, # 納特·帕格
  '10347': {'h': 0, 'w': 1}, # 老瞎眼
  '10351': {'h': 0, 'w': 0}, # 有耐心的刺客
  '10354': {'h': 0, 'w': 0}, # 深淵領主
  '10361': {'h': 1, 'w': 2}, # 準備
  '10363': {'h': 0, 'w': 1}, # 預言者費倫
  '10364': {'h': 0, 'w': 1}, # 炎爆術
  '10365': {'h': 1, 'w': 0}, # 解任務的冒險者
  '10385': {'h': 1, 'w': 0}, # 兇蠻
  '10389': {'h': 0, 'w': 2}, # 海巨人
  '10400': {'h': 0, 'w': 2}, # 暗影形態
  '10407': {'h': 1, 'w': 2}, # 盾牌猛擊
  '10418': {'h': 0, 'w': 2}, # 靈魂虹吸
  '10428': {'h': 1, 'w': 0}, # 南海船長
  '10430': {'h': 0, 'w': 2}, # 法術扭曲者
  '10440': {'h': 1, 'w': 2}, # 星殞術
  '10457': {'h': 1, 'w': 2}, # 日行者
  '10459': {'h': 1, 'w': 0}, # 正義之劍
  '10460': {'h': 0, 'w': 1}, # 希瓦娜斯·風行者
  '10463': {'h': 0, 'w': 0}, # 比斯巨獸
  '10464': {'h': 0, 'w': 1}, # 黑騎士
  '10482': {'h': 1, 'w': 0}, # 扭曲虛空
  '10485': {'h': 0, 'w': 2}, # 升級!
  '10490': {'h': 0, 'w': 2}, # 氣化
  '10521': {'h': 0, 'w': 1}, # 伊瑟拉
  '10523': {'h': 0, 'w': 0}, # 精英牛頭大佬
  # GVG
  '12005': {'h': 1, 'w': 2}, # 麥迪文的回音
  '12008': {'h': 1, 'w': 2}, # 聖光炸彈
  '12014': {'h': 0, 'w': 1}, # 沃金
  '12016': {'h': 0, 'w': 2}, # 惡魔劫奪者
  '12019': {'h': 0, 'w': 2}, # 惡魔之心
  '12021': {'h': 0, 'w': 1}, # 瑪爾加尼斯
  '12025': {'h': 0, 'w': 2}, # 齒輪大師的扳手
  '12026': {'h': 1, 'w': 0}, # 獨眼騙子
  '12027': {'h': 0, 'w': 2}, # 假死
  '12029': {'h': 0, 'w': 1}, # 貿易親王加里維克斯
  '12031': {'h': 0, 'w': 2}, # 先祖之喚
  '12039': {'h': 0, 'w': 2}, # 生命之樹
  '12041': {'h': 0, 'w': 1}, # 瑪洛尼
  '12042': {'h': 1, 'w': 2}, # 動力戰錘
  '12045': {'h': 1, 'w': 2}, # 活力圖騰
  '12047': {'h': 0, 'w': 2}, # 黑暗幽光
  '12050': {'h': 0, 'w': 1}, # 奈普圖隆
  '12055': {'h': 1, 'w': 2}, # 萬獸之王
  '12056': {'h': 0, 'w': 0}, # 破壞工作
  '12058': {'h': 0, 'w': 1}, # 加茲瑞拉
  '12059': {'h': 0, 'w': 2}, # 彈跳鋒刃
  '12061': {'h': 0, 'w': 0}, # 粉碎
  '12065': {'h': 0, 'w': 0}, # 鋼鐵破滅邪神
  '12069': {'h': 0, 'w': 0}, # 齒輪巨錘
  '12070': {'h': 0, 'w': 2}, # 軍需官
  '12073': {'h': 0, 'w': 0}, # 伯瓦爾·弗塔根
  '12076': {'h': 1, 'w': 0}, # 砂槌薩滿
  '12087': {'h': 0, 'w': 0}, # 憎惡魔像
  '12097': {'h': 1, 'w': 0}, # 攻城機具
  '12098': {'h': 1, 'w': 0}, # 熱砂狙擊手
  '12101': {'h': 1, 'w': 0}, # 更瘋狂的炸彈客
  '12103': {'h': 1, 'w': 0}, # 地精實驗家
  '12107': {'h': 1, 'w': 0}, # 哥布林工兵
  '12109': {'h': 1, 'w': 0}, # 小小驅魔者
  '12116': {'h': 1, 'w': 2}, # 大哥布林
  '12117': {'h': 0, 'w': 0}, # 有駕駛的飛天魔像
  '12118': {'h': 1, 'w': 0}, # 拾荒機器人
  '12119': {'h': 0, 'w': 2}, # 強化機器人
  '12120': {'h': 0, 'w': 0}, # 重組轉化師
  '12122': {'h': 0, 'w': 1}, # 爆爆博士
  '12124': {'h': 0, 'w': 0}, # 彌米倫之首
  '12126': {'h': 0, 'w': 0}, # 巨魔莫古
  '12127': {'h': 0, 'w': 1}, # 敵人收割者4000
  '12128': {'h': 0, 'w': 1}, # 斯尼德的伐木機
  '12129': {'h': 0, 'w': 0}, # 托斯利
  '12131': {'h': 0, 'w': 0}, # 加茲魯維
  '12132': {'h': 0, 'w': 1}, # 特洛格佐爾
  '12133': {'h': 0, 'w': 0}, # 布靈登3000型
  '12135': {'h': 1, 'w': 0}, # 發條巨人
  '12136': {'h': 0, 'w': 0}, # 嬌小的法術干擾師
}

token = [
  '11032', '11033', '11034', '11035', '11036', '11037', '11038', '11039', '11040', '11041', '11042', '11043', '11044', '11045', '11046', '11047', '11048', '11049', '11050', '11051', '11052', '11053', '11054', '11055', '11056', '11057', '11058', '11059', '11060', '11061', '11062', '11063', '11064'
  , '11065', '11066', '11067', '11068', '11069', '11070', '11071', '11072', '11073', '11074', '11075', '11076', '11077', '11078', '11079', '11080', '11081', '11082', '11083', '11084', '11085', '12030', '12033', '12034', '12037', '12038', '12048', '12049', '12054', '12066', '12091', '12104', '12123'
  , '12138', '12139', '12140', '12141', '12142', '12143', '12144', '13005', '13006', '13007', '13012', '13016', '13017', '13021', '13022', '13023', '13024', '13029', '13030', '13037', '13038', '13043', '13048', '13053', '13054', '13057', '13058', '13061', '13062', '13065', '13066', '13067', '13072'
  , '13073', '13074', '13075', '13076', '13081', '13087', '13088', '13089', '13090', '13091', '13092', '13093', '13094', '13095', '13096', '13097', '13098', '13108', '13109', '13110', '13111', '13115', '13116', '13120', '13123', '13124', '13127', '13128', '13131', '13132', '13137', '13138', '13143'
  , '13144', '13145', '13150', '13153', '13154', '13155', '13158', '13159', '13166', '13169', '13174', '13175', '13176', '13177', '13179', '13190', '13201', '10524', '10525', '10526', '10527', '10528', '10529', '10522', '10514', '10515', '10501', '10492', '10486', '10476', '10477', '10478', '10465'
  , '10454', '10439', '10433', '10437', '10431', '10422', '10381', '10336', '10340', '10342', '10320', '10296', '10294', '10264', '10267', '10241', '10239', '10233', '10174', '10152', '10147', '10134', '10135', '10136', '10117', '10119', '10441', '10017', '10018', '10037', '10045', '10047', '10052'
  , '10054', '10065', '10083', '10086', '10196', '10214', '10350', '10436', '10108', '10115', '10122', '10146', '10227', '10357', '10375', '10513', '10167', '12125', '10397', '10402', '10124'
]

have = [
  '10001', '10295', '12023', '10293', '10292', '10109', '10455', '10399', '10105', '10104', '10101', '10059', '10058', '10057', '10056', '12102', '10053', '10359', '10200', '12060', '12062', '12063', '12064', '10194', '10190', '10423', '12100', '10518', '10199', '10519', '12114', '10456', '12079'
  , '12078', '12012', '10438', '12022', '10123', '10121', '10120', '10127', '10125', '12018', '12015', '10516', '12017', '12011', '12013', '12071', '10401', '10337', '10334', '12077', '12086', '12084', '12085', '12082', '12083', '12080', '12081', '12121', '12088', '12089', '10143', '10142', '10411'
  , '10149', '10148', '10387', '10386', '10025', '10023', '10021', '10421', '10426', '12036', '12035', '10424', '10425', '10318', '10163', '10313', '10248', '10240', '10244', '10245', '10093', '10092', '10091', '10096', '10095', '10094', '12108', '12052', '10004', '10007', '10002', '10164', '12051'
  , '12053', '12040', '10369', '12057', '10366', '10362', '10352', '12067', '10353', '10481', '10483', '10484', '10260', '10261', '10265', '10358', '10466', '10508', '12113', '10462', '10461', '10505', '10468', '12068', '12075', '10063', '10060', '10066', '10064', '10416', '10211', '10210', '10413'
  , '10341', '12074', '10184', '10181', '10182', '10286', '10280', '10448', '10118', '10116', '10112', '10323', '10179', '10040', '10043', '12106', '10236', '10235', '10234', '10520', '12105', '12095', '12094', '12096', '12090', '12093', '12092', '10388', '12099', '10137', '10132', '10139', '10429'
  , '10034', '10033', '12009', '12006', '12004', '12002', '12003', '12001', '12032', '12044', '10300', '10512', '10253', '10254', '10412', '12115', '10150', '10156', '10408', '10159', '10391', '10393', '12020', '10396', '10013', '10012', '10010', '12028', '10016', '10014', '10510', '12010', '12137'
  , '10494', '10493', '10491', '10277', '10274', '10271', '10279', '10278', '10452', '12110', '12111', '12112', '10432', '10088', '10434', '12072', '10071', '10072', '10079', '10172', '12043', '12046', '10376', '10373', '10370', '10006', '10080', '10113', '10202', '10344', '10348', '10367', '10470'
  , '10471', '12007', '12130', '12134'
]


def checkToken():
  for k, v in my.iteritems():
    if k in token:
      print k

def checkHave():
  for k, v in my.iteritems():
    if k in have:
      print k

    if v['h'] == 2:
      print k,
      if k in have:
        print '.....remove from "my"'
      else:
        print '.....move to "have"'


if __name__ == '__main__':
  runit()
  checkToken()
  print 'checkToken() okay'
  checkHave()
  print 'checkHave() okay'

