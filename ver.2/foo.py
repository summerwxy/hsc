#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import re
import xlsxwriter
from xml.dom import minidom

hearthstonePath = "D:\misc\Hearthstone" # end without \

# copy from https://github.com/Sembiance/hearthstonejson/blob/master/generate.js
USED_TAGS = ["CardID", "CardName", "CardSet", "CardType", "Faction", "Rarity", "Cost", "Atk", "Health", "Durability", "CardTextInHand", "CardTextInPlay", "FlavorText", "ArtistName", "Collectible", "Elite", "Race", "Class", "HowToGetThisCard", "HowToGetThisGoldCard"]
IGNORED_TAGS = ["AttackVisualType", "EnchantmentBirthVisual", "EnchantmentIdleVisual", "TargetingArrowText", "DevState", "TriggerVisual", "Recall", "AIMustPlay", "InvisibleDeathrattle"]
MECHANIC_TAGS = ["Windfury", "Combo", "Secret", "Battlecry", "Deathrattle", "Taunt", "Stealth", "Spellpower", "Enrage", "Freeze", "Charge", "Overload", "Divine Shield", "Silence", "Morph", "OneTurnEffect", "Poisonous", "Aura", "AdjacentBuff", "HealTarget", "GrantCharge", "ImmuneToSpellpower", "AffectedBySpellPower", "Summoned"]
ENUMID_TO_NAME = {
	185 : "CardName",
	183 : "CardSet",
	202 : "CardType",
	201 : "Faction",
	199 : "Class",
	203 : "Rarity",
	48 : "Cost",
	251 : "AttackVisualType",
	184 : "CardTextInHand",
	47 : "Atk",
	45 : "Health",
	321 : "Collectible",
	342 : "ArtistName",
	351 : "FlavorText",
	32 : "TriggerVisual",
	330 : "EnchantmentBirthVisual",
	331 : "EnchantmentIdleVisual",
	268 : "DevState",
	365 : "HowToGetThisGoldCard",
	190 : "Taunt",
	364 : "HowToGetThisCard",
	338 : "OneTurnEffect",
	293 : "Morph",
	208 : "Freeze",
	252 : "CardTextInPlay",
	325 : "TargetingArrowText",
	189 : "Windfury",
	218 : "Battlecry",
	200 : "Race",
	192 : "Spellpower",
	187 : "Durability",
	197 : "Charge",
	362 : "Aura",
	361 : "HealTarget",
	349 : "ImmuneToSpellpower",
	194 : "Divine Shield",
	350 : "AdjacentBuff",
	217 : "Deathrattle",
	191 : "Stealth",
	220 : "Combo",
	339 : "Silence",
	212 : "Enrage",
	370 : "AffectedBySpellPower",
	240 : "Cant Be Damaged",
	114 : "Elite",
	219 : "Secret",
	363 : "Poisonous",
	215 : "Recall",
	340 : "Counter",
	205 : "Summoned",
	367 : "AIMustPlay",
	335 : "InvisibleDeathrattle",
	377 : "UKNOWN_HasOnDrawEffect",
	388 : "SparePart",
	389 : "UNKNOWN_DuneMaulShaman",
	380 : "UNKNOWN_Blackrock_Heroes",
	402 : "UNKNOWN_Intense_Gaze",
	401 : "UNKNOWN_BroodAffliction"
}

BOOLEAN_TYPES = ["Collectible", "Elite"];
IGNORED_TAG_NAMES = ["text", "MasterPower", "Power", "TriggeredPowerHistoryInfo", "EntourageCard"]

TAG_VALUE_MAPS = {
	"CardSet" : {
    0 : None,
		2 : "Basic",
		3 : "Classic",
		4 : "Reward",
		5 : "Missions",
		7 : "System",
		8 : "Debug",
		11 : "Promotion",
		12 : "Curse of Naxxramas",
		13 : "Goblins vs Gnomes",
		14 : "Blackrock Mountain",
    15 : "The Grand Tournament",
    # new cards edit here
		16 : "Credits",
		17 : "Hero Skins",
		18 : "Tavern Brawl"
	}, "CardType" : {
    0 : None,
		3 : "Hero",
		4 : "Minion",
		5 : "Spell",
		6 : "Enchantment",
		7 : "Weapon",
		10 : "Hero Power"
	}, "Faction" : {
    0 : None,
		1 : "Horde",
		2 : "Alliance",
		3 : "Neutral"
	}, "Rarity" : {
		0 : None,
		1 : "Common",
		2 : "Free",
		3 : "Rare",
		4 : "Epic",
		5 : "Legendary"
	}, "Race" : {
    0 : None,
		14 : "Murloc",
		15 : "Demon",
		20 : "Beast",
		21 : "Totem",
		23 : "Pirate",
		24 : "Dragon",
		17 : "Mech"
	}, "Class" : {
		0 : None,
		2 : "Druid",
		3 : "Hunter",
		4 : "Mage",
		5 : "Paladin",
		6 : "Priest",
		7 : "Rogue",
		8 : "Shaman",
		9 : "Warlock",
		10 : "Warrior",
		11 : "Dream"
	}
}

token = ['FP1_006', 'CS2_050', 'CS2_051', 'CS2_052', 'CS2_082', 'CS2_boar', 'CS2_mirror', 'CS2_tk1', 'GAME_002', 'GAME_005', 'GAME_006', 'hexfrog', 'NEW1_009', 'NEW1_032', 'NEW1_033', 'NEW1_034', 'skele11', 'PlaceholderCard', 'CS2_152', 'ds1_whelptoken', 'EX1_165t1', 'EX1_165t2', 'EX1_323w', 'EX1_tk11', 'EX1_tk28', 'EX1_tk29', 'EX1_tk34', 'EX1_tk9', 'skele21', 'Mekka1', 'Mekka2', 'Mekka3', 'Mekka4', 'Mekka4t', 'PRO_001at', 'EX1_finkle', 'EX1_598', 'AT_132_ROGUEt', 'AT_132_SHAMANa', 'AT_132_SHAMANb', 'AT_132_SHAMANc', 'AT_132_SHAMANd']

# in pack
pack = ['CS1_069', 'CS1_129', 'CS2_028', 'CS2_031', 'CS2_038', 'CS2_059', 'CS2_073', 'CS2_104', 'CS2_117', 'CS2_146', 'CS2_151', 'CS2_161', 'CS2_169', 'CS2_188', 'CS2_203', 'CS2_221', 'CS2_227', 'CS2_231', 'CS2_233', 'DS1_188', 'EX1_001', 'EX1_004', 'EX1_005', 'EX1_006', 'EX1_007', 'EX1_008', 'EX1_009', 'EX1_010', 'EX1_017', 'EX1_020', 'EX1_021', 'EX1_023', 'EX1_028', 'EX1_029', 'EX1_033', 'EX1_043', 'EX1_046', 'EX1_048', 'EX1_049', 'EX1_050', 'EX1_055', 'EX1_057'
    , 'EX1_058', 'EX1_059', 'EX1_076', 'EX1_080', 'EX1_082', 'EX1_083', 'EX1_089', 'EX1_091', 'EX1_093', 'EX1_096', 'EX1_097', 'EX1_102', 'EX1_103', 'EX1_110', 'EX1_124', 'EX1_126', 'EX1_128', 'EX1_130', 'EX1_131', 'EX1_132', 'EX1_133', 'EX1_134', 'EX1_136', 'EX1_137', 'EX1_144', 'EX1_154', 'EX1_155', 'EX1_158', 'EX1_160', 'EX1_161', 'EX1_162', 'EX1_164', 'EX1_165', 'EX1_166', 'EX1_170', 'EX1_178', 'EX1_238', 'EX1_241', 'EX1_243', 'EX1_245', 'EX1_247', 'EX1_248'
    , 'EX1_251', 'EX1_258', 'EX1_274', 'EX1_275', 'EX1_283', 'EX1_284', 'EX1_289', 'EX1_294', 'EX1_295', '', 'X1_298', 'EX1_301', 'EX1_303', 'EX1_304', 'EX1_310', 'EX1_315', 'EX1_316', 'EX1_317', 'EX1_319', 'EX1_332', 'EX1_334', 'EX1_335', 'EX1_339', 'EX1_341', 'EX1_349', 'EX1_355', 'EX1_362', 'EX1_363', 'EX1_365', 'EX1_379', 'EX1_383', 'EX1_390', 'EX1_391', 'EX1_392', 'EX1_393', 'EX1_396', 'EX1_398', 'EX1_405', 'EX1_408', 'EX1_411', 'EX1_412', 'EX1_414', 'EX1_507'
    , 'EX1_531', 'EX1_533', 'EX1_534', 'EX1_536', 'EX1_538', 'EX1_544', 'EX1_554', 'EX1_556', 'EX1_560', 'EX1_562', 'EX1_567', 'EX1_570', 'EX1_578', 'EX1_583', 'EX1_584', 'EX1_591', 'EX1_595', 'EX1_596', 'EX1_597', 'EX1_603', 'EX1_604', 'EX1_607', 'EX1_608', 'EX1_609', 'EX1_610', 'EX1_611', 'EX1_612', 'EX1_614', 'EX1_616', 'EX1_617', 'EX1_619', 'EX1_621', 'EX1_623', 'EX1_624', 'NEW1_010', 'NEW1_012', 'NEW1_014', 'NEW1_018', 'NEW1_019', 'NEW1_020', 'NEW1_022'
    , 'NEW1_023', 'NEW1_025', 'NEW1_026', 'NEW1_027', 'NEW1_030', 'NEW1_036', 'NEW1_041', 'tt_004', 'EX1_062', 'GVG_001', 'GVG_002', 'GVG_003', 'GVG_004', 'GVG_006', 'GVG_007', 'GVG_009', 'GVG_010', 'GVG_011', 'GVG_012', 'GVG_013', 'GVG_015', 'GVG_017', 'GVG_018', 'GVG_020', 'GVG_022', 'GVG_023', 'GVG_027', 'GVG_030', 'GVG_031', 'GVG_032', 'GVG_034', 'GVG_037', 'GVG_038', 'GVG_040', 'GVG_043', 'GVG_044', 'GVG_045', 'GVG_048', 'GVG_051', 'GVG_053', 'GVG_054', 'GVG_055'
    , 'GVG_057', 'GVG_058', 'GVG_061', 'GVG_062', 'GVG_064', 'GVG_065', 'GVG_067', 'GVG_068', 'GVG_069', 'GVG_070', 'GVG_071', 'GVG_072', 'GVG_073', 'GVG_074', 'GVG_075', 'GVG_076', 'GVG_078', 'GVG_079', 'GVG_080', 'GVG_081', 'GVG_082', 'GVG_083', 'GVG_084', 'GVG_085', 'GVG_088', 'GVG_089', 'GVG_091', 'GVG_093', 'GVG_094', 'GVG_096', 'GVG_098', 'GVG_099', 'GVG_100', 'GVG_101', 'GVG_102', 'GVG_103', 'GVG_109', 'GVG_116', 'GVG_120', 'GVG_123', 'EX1_298', 'EX1_410'
    , 'NEW1_037', 'AT_100', 'AT_074', 'AT_089', 'AT_103', 'AT_114', 'AT_082', 'AT_001', 'AT_005', 'AT_006', 'AT_020', 'AT_021', 'AT_022', 'AT_024', 'AT_028', 'AT_046', 'AT_049', 'AT_052', 'AT_053', 'AT_055', 'AT_058', 'AT_059', 'AT_060', 'AT_064', 'AT_065', 'AT_068', 'AT_070', 'AT_083', 'AT_085', 'AT_086', 'AT_087', 'AT_088', 'AT_090', 'AT_091', 'AT_092', 'AT_093', 'AT_095', 'AT_106', 'AT_111', 'AT_112'
    , 'AT_133'
    ]

my = {
  'CS2_181': {'h': 1, 'w': 0}, # 受傷的大劍師
  'EX1_032': {'h': 1, 'w': 0}, # 日行者
  'EX1_044': {'h': 1, 'w': 0}, # 解任務的冒險者
  'EX1_045': {'h': 1, 'w': 0}, # 遠古看守者
  'EX1_067': {'h': 1, 'w': 0}, # 銀白指揮官
  'EX1_085': {'h': 1, 'w': 0}, # 精神控制技師
  'EX1_105': {'h': 1, 'w': 0}, # 山嶺巨人
  'EX1_145': {'h': 1, 'w': 0}, # 準備
  'EX1_279': {'h': 1, 'w': 0}, # 炎爆術
  'EX1_287': {'h': 1, 'w': 0}, # 法術反制
  'EX1_309': {'h': 1, 'w': 0}, # 靈魂虹吸
  'EX1_312': {'h': 1, 'w': 0}, # 扭曲虛空
  'EX1_354': {'h': 1, 'w': 0}, # 聖療術
  'EX1_366': {'h': 1, 'w': 0}, # 正義之劍
  'EX1_382': {'h': 1, 'w': 0}, # 奧多爾保安官
  'EX1_407': {'h': 1, 'w': 0}, # 鬥毆
  'EX1_509': {'h': 1, 'w': 0}, # 魚人招潮者
  'EX1_522': {'h': 1, 'w': 0}, # 有耐心的刺客
  'EX1_537': {'h': 1, 'w': 0}, # 爆裂射擊
  'EX1_564': {'h': 1, 'w': 0}, # 無面操縱者
  'EX1_571': {'h': 1, 'w': 0}, # 自然之力
  'EX1_590': {'h': 1, 'w': 0}, # 血騎士
  'NEW1_007': {'h': 1, 'w': 0}, # 星殞術
  'NEW1_017': {'h': 1, 'w': 0}, # 飢餓的螃蟹
  'NEW1_021': {'h': 1, 'w': 0}, # 末日預言者
  'GVG_005': {'h': 1, 'w': 0}, # 麥迪文的回音
  'GVG_008': {'h': 1, 'w': 0}, # 聖光炸彈
  'GVG_016': {'h': 1, 'w': 0}, # 惡魔劫奪者
  'GVG_025': {'h': 1, 'w': 0}, # 獨眼騙子
  'GVG_036': {'h': 1, 'w': 0}, # 動力戰錘
  'GVG_039': {'h': 1, 'w': 0}, # 活力圖騰
  'GVG_046': {'h': 1, 'w': 0}, # 萬獸之王
  'GVG_066': {'h': 1, 'w': 0}, # 砂槌薩滿
  'GVG_086': {'h': 1, 'w': 0}, # 攻城機具
  'GVG_087': {'h': 1, 'w': 0}, # 熱砂狙擊手
  'GVG_090': {'h': 1, 'w': 0}, # 更瘋狂的炸彈客
  'GVG_092': {'h': 1, 'w': 0}, # 地精實驗家
  'GVG_095': {'h': 1, 'w': 0}, # 哥布林工兵
  'GVG_097': {'h': 1, 'w': 0}, # 小小驅魔者
  'GVG_104': {'h': 1, 'w': 0}, # 大哥布林
  'GVG_106': {'h': 1, 'w': 0}, # 拾荒機器人
  'GVG_121': {'h': 1, 'w': 0}, # 發條巨人
  'CS2_053': {'h': 0, 'w': 0}, # 視界術
  'EX1_002': {'h': 0, 'w': 0}, # 黑騎士
  'EX1_012': {'h': 0, 'w': 0}, # 血法師薩爾諾斯
  'EX1_014': {'h': 0, 'w': 0}, # 穆克拉
  'EX1_016': {'h': 0, 'w': 0}, # 希瓦娜斯·風行者
  'EX1_095': {'h': 1, 'w': 0}, # 加基森拍賣師
  'EX1_100': {'h': 0, 'w': 0}, # 博學行者阿洲
  'EX1_116': {'h': 0, 'w': 0}, # 炸雞勇者
  'EX1_249': {'h': 0, 'w': 0}, # 迦頓男爵
  'EX1_250': {'h': 0, 'w': 0}, # 土元素
  'EX1_259': {'h': 0, 'w': 0}, # 閃電風暴
  'EX1_313': {'h': 0, 'w': 0}, # 深淵領主
  'EX1_320': {'h': 0, 'w': 0}, # 末日災厄
  'EX1_323': {'h': 0, 'w': 0}, # 賈拉克瑟斯領主
  'EX1_345': {'h': 0, 'w': 0}, # 心理遊戲
  'EX1_350': {'h': 0, 'w': 0}, # 預言者費倫
  'EX1_384': {'h': 0, 'w': 0}, # 復仇之怒
  'EX1_402': {'h': 0, 'w': 0}, # 護甲鍛造師
  'EX1_409': {'h': 0, 'w': 0}, # 升級!
  'EX1_543': {'h': 0, 'w': 0}, # 克洛許王
  'EX1_549': {'h': 0, 'w': 0}, # 狂野怒火
  'EX1_557': {'h': 0, 'w': 0}, # 納特·帕格
  'EX1_558': {'h': 0, 'w': 0}, # 哈里遜·瓊斯
  'EX1_559': {'h': 0, 'w': 0}, # 大法師安東尼達斯
  'EX1_561': {'h': 0, 'w': 0}, # 雅立史卓莎
  'EX1_563': {'h': 0, 'w': 0}, # 瑪里苟斯
  'EX1_572': {'h': 0, 'w': 0}, # 伊瑟拉
  'EX1_573': {'h': 0, 'w': 0}, # 塞納留斯
  'EX1_575': {'h': 0, 'w': 0}, # 法力之潮圖騰
  'EX1_577': {'h': 0, 'w': 0}, # 比斯巨獸
  'EX1_586': {'h': 0, 'w': 0}, # 海巨人
  'EX1_594': {'h': 1, 'w': 0}, # 氣化
  'EX1_613': {'h': 0, 'w': 0}, # 艾德溫·范克里夫
  'EX1_620': {'h': 0, 'w': 0}, # 熔火巨人
  'EX1_625': {'h': 0, 'w': 0}, # 暗影形態
  'EX1_626': {'h': 0, 'w': 0}, # 群體驅魔
  'NEW1_005': {'h': 1, 'w': 0}, # 綁匪
  'NEW1_008': {'h': 0, 'w': 0}, # 知識古樹
  'NEW1_024': {'h': 0, 'w': 0}, # 綠皮船長
  'NEW1_029': {'h': 0, 'w': 0}, # 米歐浩斯·曼納斯頓
  'NEW1_038': {'h': 0, 'w': 0}, # 戈魯爾
  'NEW1_040': {'h': 0, 'w': 0}, # 霍格
  'tt_010': {'h': 0, 'w': 0}, # 法術扭曲者
  'NEW1_016': {'h': 0, 'w': 0}, # 船長的鸚鵡
  'EX1_112': {'h': 0, 'w': 0}, # 傑爾賓·梅卡托克
  'PRO_001': {'h': 0, 'w': 0}, # 精英牛頭大佬
  'GVG_014': {'h': 0, 'w': 0}, # 沃金
  'GVG_019': {'h': 0, 'w': 0}, # 惡魔之心
  'GVG_021': {'h': 0, 'w': 0}, # 瑪爾加尼斯
  'GVG_024': {'h': 0, 'w': 0}, # 齒輪大師的扳手
  'GVG_026': {'h': 0, 'w': 0}, # 假死
  'GVG_028': {'h': 0, 'w': 0}, # 貿易親王加里維克斯
  'GVG_029': {'h': 0, 'w': 0}, # 先祖之喚
  'GVG_033': {'h': 0, 'w': 0}, # 生命之樹
  'GVG_035': {'h': 0, 'w': 0}, # 瑪洛尼
  'GVG_041': {'h': 0, 'w': 0}, # 黑暗幽光
  'GVG_042': {'h': 0, 'w': 0}, # 奈普圖隆
  'GVG_047': {'h': 0, 'w': 0}, # 破壞工作
  'GVG_049': {'h': 0, 'w': 0}, # 加茲瑞拉
  'GVG_050': {'h': 0, 'w': 0}, # 彈跳鋒刃
  'GVG_052': {'h': 0, 'w': 0}, # 粉碎
  'GVG_056': {'h': 0, 'w': 0}, # 鋼鐵破滅邪神
  'GVG_059': {'h': 0, 'w': 0}, # 齒輪巨錘
  'GVG_060': {'h': 0, 'w': 0}, # 軍需官
  'GVG_063': {'h': 0, 'w': 0}, # 伯瓦爾·弗塔根
  'GVG_077': {'h': 0, 'w': 0}, # 憎惡魔像
  'GVG_105': {'h': 0, 'w': 0}, # 有駕駛的飛天魔像
  'GVG_107': {'h': 0, 'w': 0}, # 強化機器人
  'GVG_108': {'h': 0, 'w': 0}, # 重組轉化師
  'GVG_110': {'h': 0, 'w': 0}, # 爆爆博士
  'GVG_111': {'h': 0, 'w': 0}, # 彌米倫之首
  'GVG_112': {'h': 0, 'w': 0}, # 巨魔莫古
  'GVG_113': {'h': 0, 'w': 0}, # 敵人收割者4000
  'GVG_114': {'h': 0, 'w': 0}, # 斯尼德的伐木機
  'GVG_115': {'h': 0, 'w': 0}, # 托斯利
  'GVG_117': {'h': 0, 'w': 0}, # 加茲魯維
  'GVG_118': {'h': 0, 'w': 0}, # 特洛格佐爾
  'GVG_119': {'h': 0, 'w': 0}, # 布靈登3000型
  'GVG_122': {'h': 0, 'w': 0}, # 嬌小的法術干擾師   
  # TGT
  'AT_002': {'h': 0, 'w': 0}, # 火焰稻草人
  'AT_003': {'h': 1, 'w': 0}, # 陣亡英雄之靈
  'AT_004': {'h': 0, 'w': 0}, # 秘法衝擊
  'AT_007': {'h': 0, 'w': 0}, # 魔法鏢客
  'AT_008': {'h': 0, 'w': 0}, # 凜懼島飛龍
  'AT_009': {'h': 0, 'w': 0}, # 羅甯
  'AT_010': {'h': 0, 'w': 0}, # 山羊牧人
  'AT_011': {'h': 0, 'w': 0}, # 神聖勇士
  'AT_012': {'h': 1, 'w': 0}, # 暗影爪牙
  'AT_013': {'h': 1, 'w': 0}, # 真言術：耀
  'AT_014': {'h': 1, 'w': 0}, # 暗影惡魔
  'AT_015': {'h': 0, 'w': 0}, # 歸順
  'AT_016': {'h': 0, 'w': 0}, # 混亂
  'AT_017': {'h': 1, 'w': 0}, # 暮光守護者
  'AT_018': {'h': 0, 'w': 0}, # 告解者帕爾璀絲
  'AT_019': {'h': 0, 'w': 0}, # 恐懼戰馬
  'AT_023': {'h': 1, 'w': 0}, # 虛無粉碎者
  'AT_025': {'h': 1, 'w': 0}, # 黑暗交易
  'AT_026': {'h': 1, 'w': 0}, # 憤怒守衛
  'AT_027': {'h': 0, 'w': 0}, # 威爾弗雷德‧菲斯巴恩
  'AT_029': {'h': 0, 'w': 0}, # 海賊
  'AT_030': {'h': 0, 'w': 0}, # 幽暗城驍士
  'AT_031': {'h': 0, 'w': 0}, # 扒手
  'AT_032': {'h': 1, 'w': 0}, # 黑市商人
  'AT_033': {'h': 0, 'w': 0}, # 盜竊
  'AT_034': {'h': 0, 'w': 0}, # 毒刃
  'AT_035': {'h': 0, 'w': 0}, # 地底潛伏
  'AT_036': {'h': 0, 'w': 0}, # 阿努巴拉克
  'AT_037': {'h': 1, 'w': 0}, # 糾纏之根
  'AT_038': {'h': 1, 'w': 0}, # 達納蘇斯志士
  'AT_039': {'h': 1, 'w': 0}, # 蠻荒戰鬥者
  'AT_040': {'h': 0, 'w': 0}, # 荒野行者
  'AT_041': {'h': 0, 'w': 0}, # 荒野騎士
  'AT_042': {'h': 1, 'w': 0}, # 刃牙德魯伊
  'AT_043': {'h': 0, 'w': 0}, # 星體共融
  'AT_044': {'h': 0, 'w': 0}, # 堆肥
  'AT_045': {'h': 0, 'w': 0}, # 艾維娜
  'AT_047': {'h': 0, 'w': 0}, # 德萊尼圖騰雕刻師
  'AT_048': {'h': 0, 'w': 0}, # 治療波
  'AT_050': {'h': 0, 'w': 0}, # 充能戰錘
  'AT_051': {'h': 0, 'w': 0}, # 元素毀滅
  'AT_054': {'h': 0, 'w': 0}, # 喚霧者
  'AT_056': {'h': 0, 'w': 0}, # 強力射擊
  'AT_057': {'h': 0, 'w': 0}, # 獸欄管理員
  'AT_061': {'h': 0, 'w': 0}, # 全面備戰
  'AT_062': {'h': 0, 'w': 0}, # 蜘蛛囊
  'AT_063': {'h': 0, 'w': 0}, # 酸喉
  'AT_066': {'h': 1, 'w': 0}, # 奧格瑪志士
  'AT_067': {'h': 0, 'w': 0}, # 猛瑪象人首領
  'AT_069': {'h': 0, 'w': 0}, # 練習夥伴
  'AT_071': {'h': 0, 'w': 0}, # 雅立史卓莎的勇士
  'AT_072': {'h': 0, 'w': 0}, # 瓦里安‧烏瑞恩
  'AT_073': {'h': 0, 'w': 0}, # 運動精神
  'AT_075': {'h': 1, 'w': 0}, # 戰馬訓練師
  'AT_076': {'h': 0, 'w': 0}, # 魚人騎士
  'AT_077': {'h': 0, 'w': 0}, # 銀白長槍
  'AT_078': {'h': 0, 'w': 0}, # 高手過招
  'AT_079': {'h': 0, 'w': 0}, # 神秘挑戰者
  'AT_080': {'h': 0, 'w': 0}, # 要塞指揮官
  'AT_081': {'h': 0, 'w': 0}, # 『純淨者』埃卓克
  'AT_084': {'h': 0, 'w': 0}, # 槍僮
  'AT_094': {'h': 1, 'w': 0}, # 火焰雜耍師
  'AT_096': {'h': 0, 'w': 0}, # 發條騎士
  'AT_097': {'h': 0, 'w': 0}, # 聯賽觀眾
  'AT_098': {'h': 0, 'w': 0}, # 雜耍吞法者
  'AT_099': {'h': 0, 'w': 0}, # 科多獸騎士
  'AT_101': {'h': 0, 'w': 0}, # 鬥技場鬥士
  'AT_102': {'h': 1, 'w': 0}, # 捕獲的蟄猛巨蟲
  'AT_104': {'h': 1, 'w': 0}, # 巨牙矛騎兵
  'AT_105': {'h': 0, 'w': 0}, # 受傷的科瓦迪爾
  'AT_108': {'h': 1, 'w': 0}, # 裝甲戰馬
  'AT_109': {'h': 1, 'w': 0}, # 銀白巡邏兵
  'AT_110': {'h': 0, 'w': 0}, # 大競技場經理
  'AT_113': {'h': 1, 'w': 0}, # 募兵官
  'AT_115': {'h': 1, 'w': 0}, # 擊劍教練
  'AT_116': {'h': 1, 'w': 0}, # 龍眠使者
  'AT_117': {'h': 0, 'w': 0}, # 大會主持人
  'AT_118': {'h': 0, 'w': 0}, # 大十字軍
  'AT_119': {'h': 0, 'w': 0}, # 科瓦迪爾劫掠者
  'AT_120': {'h': 0, 'w': 0}, # 冰霜巨人
  'AT_121': {'h': 0, 'w': 0}, # 大明星
  'AT_122': {'h': 0, 'w': 0}, # 『穿刺者』戈莫克
  'AT_123': {'h': 0, 'w': 0}, # 寒冽之喉
  'AT_124': {'h': 0, 'w': 0}, # 波爾夫‧拉姆榭
  'AT_125': {'h': 0, 'w': 0}, # 冰嚎
  'AT_127': {'h': 0, 'w': 0}, # 奈薩斯勇士薩拉德
  'AT_128': {'h': 0, 'w': 0}, # 骷髏騎士
  'AT_129': {'h': 0, 'w': 0}, # 菲歐拉‧光寂
  'AT_130': {'h': 0, 'w': 0}, # 海劫者
  'AT_131': {'h': 0, 'w': 0}, # 艾狄絲‧暗寂
  'AT_132': {'h': 0, 'w': 0}, # 審判者瑪瑞爾
  }

cmap = {
  'Basic': '基本',
  'Blackrock Mountain': 'BRM',
  'Classic': '經典',
  'Curse of Naxxramas': 'Naxx',
  'Goblins vs Gnomes': 'GVG',
  'The Grand Tournament': 'TGT',
  # new cards edit here
  'Minion': '手下',
  'Spell': '法術',
  'Weapon': '武器',
  'Legendary': '傳說',
  'Epic': '史詩',
  'Rare': '精良',
  'Common': '基本',
  'Free': '免費',
  'Druid': '德魯伊',
  'Hunter': '獵人',
  'Mage': '法師',
  'Paladin': '聖騎',
  'Priest': '牧師',
  'Rogue': '盜賊',
  'Shaman': '薩滿',
  'Warlock': '術士',
  'Warrior': '戰士',
  'Beast': '野獸',
  'Demon': '惡魔',
  'Dragon': '龍類',
  'Mech': '機械',
  'Murloc': '魚人',
  'Pirate': '海盜',
  'Totem': '圖騰',
    }

# extract from hearthstone
def extractFromHearthstone():
  cmd = ".\disunity_v0.3.4\disunity.bat extract %s\Data\Win\cardxml0.unity3d" % (hearthstonePath)
  os.system(cmd)

# parse *.txt to xlsx
def parseTxtFiles():
  txtPath = "%s\Data\Win\cardxml0\CAB-cardxml0\TextAsset" % (hearthstonePath) 
  doc = minidom.parse("%s\zhTW.txt" % (txtPath))
  root = doc.documentElement
  cards = root.getElementsByTagName("Entity")
  
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
  
  title = ['dust', 'total', 'total', 'have', 'have', 'need', 'need', 'wish', 'wish', 'id', 'name', 'mp', 'atk', 'hp', 'set', 'type', 'rarity', 'class', 'race', 'effect', 'desc', '', 'get', 'getcold']
  for i in range(len(title)):
    worksheet.write(0, i, title[i])

  i = 1
  for card in cards:
    cid = card.getAttribute('CardID')
    cset = getTagValue(card, "CardSet")
    ctype = getTagValue(card, "CardType")
    crarity = getTagValue(card, "Rarity")
    cclass = getTagValue(card, "Class")
    crace = getTagValue(card, "Race")

    dust = 0
    total = 2
    have = 0
    if crarity == 'Legendary':
      dust = 1600
      total = 1
    elif crarity == 'Epic':
      dust = 400
    elif crarity == 'Rare':
      dust = 100
    elif crarity == 'Common':
      dust = 40

    # about color
    format = white
    if ctype in ['Enchantment', 'Hero Power', 'Hero'] \
        or cset in ['Credits', 'Tavern Brawl', 'Missions', 'Debug'] \
        or cclass in ['Dream'] \
        or cid in token \
        or re.match(".+_\d{3}t$", cid) \
        or re.match(".+_\d{3}t2$", cid) \
        or re.match(".+_\d{3}a$", cid) \
        or re.match(".+_\d{3}b$", cid) \
        or re.match(".+_\d{3}c$", cid) \
        or cid.startswith("PART_") \
        or cid.startswith("NAX") \
        or cid.startswith("BRMA"):
      format = black
      dust = ''
      total = ''
      have = ''
    elif cset == 'Basic':
      format = basic
      dust = ''
      have = total
    elif cid in pack:
      format = green
      have = total
    elif cset in ['Curse of Naxxramas', 'Blackrock Mountain']:
      format = green
      dust = ''
      have = total
    elif crarity == 'Legendary':
      format = orange
    elif my.get(cid) and my.get(cid).get('h') == 1:
      format = blue
      have = 1
    elif my.get(cid) and my.get(cid).get('h') == 0:
      format = red


    worksheet.write(i, 0, dust)
    worksheet.write(i, 1, total)
    worksheet.write(i, 2, total and total * dust)
    worksheet.write(i, 3, have)
    worksheet.write(i, 4, have and have * dust)
    need =(total == '' and [''] or [int(total) - int(have)])[0]
    worksheet.write(i, 5, need)
    worksheet.write(i, 6, (dust == '' and [''] or [dust * need])[0])
    # TODO: wish column
    worksheet.write(i, 9, cid)
    worksheet.write(i, 10, getTagValue(card, "CardName"))
    worksheet.write(i, 11, getTagValue(card, "Cost"))
    worksheet.write(i, 12, getTagValue(card, "Atk"))
    worksheet.write(i, 13, getTagValue(card, "Health") or getTagValue(card, "Durability"))
    worksheet.write(i, 14, cmap.get(cset) or cset)
    worksheet.write(i, 15, cmap.get(ctype) or ctype)
    worksheet.write(i, 16, cmap.get(crarity) or crarity)
    worksheet.write(i, 17, cmap.get(cclass) or cclass)
    worksheet.write(i, 18, cmap.get(crace) or crace)
    worksheet.write(i, 19, remove_tags(getTagValue(card, "CardTextInHand")))
    worksheet.write(i, 20, remove_tags(getTagValue(card, "FlavorText")))
    worksheet.write(i, 21, "------")
    worksheet.write(i, 22, getTagValue(card, "HowToGetThisCard"))
    worksheet.write(i, 23, getTagValue(card, "HowToGetThisGoldCard"))
    worksheet.write(i, 24, getTagValue(card, "Faction"))
    worksheet.write(i, 25, getTagValue(card, "CardTextInPlay"))
    worksheet.write(i, 26, getTagValue(card, "ArtistName"))
    worksheet.write(i, 27, getTagValue(card, "Collectible"))
    worksheet.write(i, 28, getTagValue(card, "Elite"))

    worksheet.set_row(i, 13.5, format)

    i += 1

  ws = [5, 3, 5, 3, 5, 3, 5, 3, 5, 10, 22, 3, 3, 3, 5, 5, 5.5, 5, 5, 40, 40]
  for i in range(len(ws)):
    worksheet.set_column(i, i, ws[i])

  worksheet.autofilter('A1:AB1')
  workbook.close()


def getTagValue(card, tagName):
  result = ""
  key = [key for key, value in ENUMID_TO_NAME.items() if value == tagName][0]
  for tag in card.getElementsByTagName("Tag"):
    if tag.getAttribute("enumID") == str(key):
      if tag.getAttribute("type") == "String":
        result = tag.childNodes[0].nodeValue
      elif tag.getAttribute("type") == "":
        result = tag.getAttribute("value")
      else:
        print("wtf!!")
      break

  if TAG_VALUE_MAPS.get(tagName):
    foo = TAG_VALUE_MAPS.get(tagName)
    result = foo.get(int(result or 0))
  return result or ""

TAG_RE = re.compile(r'<[^>]+>')

def remove_tags(text):
  return TAG_RE.sub('', text)

if __name__ == '__main__':
  #extractFromHearthstone()
  parseTxtFiles()





