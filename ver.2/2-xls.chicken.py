#!/usr/bin/env python
# -*- coding: utf-8 -*-
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
pack = [
  'CS2_031', 'EX1_339', 'EX1_335', 'EX1_133', 'EX1_607', 'CS2_104', 'EX1_049', 'EX1_021', 'EX1_556', 'EX1_396', 'EX1_383', 'EX1_624', 'EX1_251', 'EX1_247'
  , 'EX1_317', 'NEW1_023', 'EX1_390', 'EX1_412', 'EX1_048', 'EX1_621', 'EX1_289', 'CS2_073', 'EX1_604', 'EX1_398', 'CS2_231', 'EX1_162', 'EX1_538'
  , 'EX1_608', 'EX1_132', 'EX1_243', 'EX1_509', 'EX1_057', 'AT_040', 'NEW1_022', 'EX1_158', 'EX1_010', 'EX1_165', 'EX1_238', 'CS2_169'
  , 'CS2_028', 'EX1_008', 'EX1_093', 'EX1_583'
  , 'GVG_100'
  , 'GVG_076'
  , 'EX1_096'
    ]

my = {
  # classic
  'CS1_069': {'h': 1, 'w': 0}, # 沼地蠕行者 536
  'CS1_129': {'h': 1, 'w': 0}, # 心靈之火 1XX
  'CS2_038': {'h': 1, 'w': 0}, # 先祖之魂 2XX
  'CS2_053': {'h': 0, 'w': 0}, # 視界術 3XX
  'CS2_059': {'h': 0, 'w': 0}, # 血之小鬼 101
  'CS2_117': {'h': 1, 'w': 0}, # 陶土議會先知 333
  'CS2_146': {'h': 0, 'w': 0}, # 南海水手 121
  'CS2_151': {'h': 1, 'w': 0}, # 白銀之手騎士 544
  'CS2_161': {'h': 1, 'w': 0}, # 拉文霍德刺客 775
  'CS2_181': {'h': 0, 'w': 0}, # 受傷的大劍師 347
  'CS2_188': {'h': 0, 'w': 0}, # 嚴厲的士官 121
  'CS2_203': {'h': 1, 'w': 0}, # 鐵喙貓頭鷹 221
  'CS2_221': {'h': 1, 'w': 0}, # 惡毒的鐵匠 546
  'CS2_227': {'h': 1, 'w': 0}, # 風險投資公司雇傭兵 576
  'CS2_233': {'h': 1, 'w': 0}, # 劍刃亂舞 2XX
  'DS1_188': {'h': 0, 'w': 0}, # 鬥士長弓 752
  'EX1_001': {'h': 0, 'w': 0}, # 護光者 112
  'EX1_002': {'h': 0, 'w': 0}, # 黑騎士 645
  'EX1_004': {'h': 0, 'w': 0}, # 年輕的女牧師 121
  'EX1_005': {'h': 0, 'w': 0}, # 王牌獵人 342
  'EX1_006': {'h': 0, 'w': 0}, # 警報機器人 303
  'EX1_007': {'h': 0, 'w': 0}, # 苦痛侍僧 313
  'EX1_009': {'h': 0, 'w': 0}, # 憤怒雞 111
  'EX1_012': {'h': 0, 'w': 0}, # 血法師薩爾諾斯 211
  'EX1_014': {'h': 0, 'w': 0}, # 穆克拉 355
  'EX1_016': {'h': 0, 'w': 0}, # 希瓦娜斯·風行者 655
  'EX1_017': {'h': 0, 'w': 0}, # 叢林豹 342
  'EX1_020': {'h': 0, 'w': 0}, # 血色十字軍 331
  'EX1_023': {'h': 1, 'w': 0}, # 銀月看守者 433
  'EX1_028': {'h': 0, 'w': 0}, # 荊棘谷猛虎 555
  'EX1_029': {'h': 0, 'w': 0}, # 麻瘋地精 121
  'EX1_032': {'h': 0, 'w': 0}, # 日行者 645
  'EX1_033': {'h': 1, 'w': 0}, # 風怒鷹身人 645
  'EX1_043': {'h': 1, 'w': 0}, # 暮光飛龍 441
  'EX1_044': {'h': 0, 'w': 0}, # 解任務的冒險者 322
  'EX1_045': {'h': 0, 'w': 0}, # 遠古看守者 245
  'EX1_046': {'h': 1, 'w': 0}, # 黑鐵矮人 444
  'EX1_050': {'h': 1, 'w': 0}, # 冷光神諭者 322
  'EX1_055': {'h': 0, 'w': 0}, # 魔法成癮者 213
  'EX1_058': {'h': 0, 'w': 0}, # 日怒保衛者 223
  'EX1_059': {'h': 0, 'w': 0}, # 瘋狂的鍊金師 222
  'EX1_067': {'h': 1, 'w': 0}, # 銀白指揮官 642
  'EX1_076': {'h': 0, 'w': 0}, # 迷你召喚師 222
  'EX1_080': {'h': 0, 'w': 0}, # 保密者 112
  'EX1_082': {'h': 1, 'w': 0}, # 瘋狂炸彈客 232
  'EX1_083': {'h': 0, 'w': 0}, # 技工大師歐沃斯巴克 333
  'EX1_085': {'h': 0, 'w': 0}, # 精神控制技師 333
  'EX1_089': {'h': 0, 'w': 0}, # 秘法魔像 342
  'EX1_091': {'h': 1, 'w': 0}, # 卡巴暗影牧師 645
  'EX1_095': {'h': 1, 'w': 0}, # 加基森拍賣師 644
  'EX1_097': {'h': 1, 'w': 0}, # 憎惡體 544
  'EX1_100': {'h': 0, 'w': 0}, # 博學行者阿洲 204
  'EX1_102': {'h': 0, 'w': 0}, # 石毀車 314
  'EX1_103': {'h': 1, 'w': 0}, # 冷光先知 323
  'EX1_105': {'h': 0, 'w': 0}, # 山嶺巨人 1288
  'EX1_110': {'h': 0, 'w': 0}, # 凱恩·血蹄 645
  'EX1_116': {'h': 0, 'w': 0}, # 炸雞勇者 562
  'EX1_124': {'h': 0, 'w': 0}, # 剔骨 2XX
  'EX1_126': {'h': 0, 'w': 0}, # 背叛 2XX
  'EX1_128': {'h': 1, 'w': 0}, # 隱蔽 1XX
  'EX1_130': {'h': 1, 'w': 0}, # 光榮犧牲 1XX
  'EX1_131': {'h': 0, 'w': 0}, # 迪菲亞頭目 222
  'EX1_134': {'h': 0, 'w': 0}, # 軍情七處密探 333
  'EX1_136': {'h': 1, 'w': 0}, # 救贖 1XX
  'EX1_137': {'h': 0, 'w': 0}, # 碎顱 3XX
  'EX1_144': {'h': 1, 'w': 0}, # 暗影閃現 0XX
  'EX1_145': {'h': 0, 'w': 0}, # 準備 0XX
  'EX1_154': {'h': 0, 'w': 0}, # 憤怒 2XX
  'EX1_155': {'h': 0, 'w': 0}, # 自然印記 3XX
  'EX1_160': {'h': 0, 'w': 0}, # 野性之力 2XX
  'EX1_161': {'h': 1, 'w': 0}, # 自然化 1XX
  'EX1_164': {'h': 0, 'w': 0}, # 滋補術 5XX
  'EX1_166': {'h': 0, 'w': 0}, # 森林看守者 424
  'EX1_170': {'h': 1, 'w': 0}, # 帝王眼鏡蛇 323
  'EX1_178': {'h': 0, 'w': 0}, # 戰爭古樹 755
  'EX1_241': {'h': 0, 'w': 0}, # 熔岩爆發 3XX
  'EX1_245': {'h': 1, 'w': 0}, # 大地震擊 1XX
  'EX1_248': {'h': 0, 'w': 0}, # 野性之魂 3XX
  'EX1_249': {'h': 0, 'w': 0}, # 迦頓男爵 775
  'EX1_250': {'h': 0, 'w': 0}, # 土元素 578
  'EX1_258': {'h': 1, 'w': 0}, # 無縛的元素 324
  'EX1_259': {'h': 0, 'w': 0}, # 閃電風暴 3XX
  'EX1_274': {'h': 0, 'w': 0}, # 以太秘法師 433
  'EX1_275': {'h': 1, 'w': 0}, # 冰錐術 4XX
  'EX1_279': {'h': 0, 'w': 0}, # 炎爆術 10XX
  'EX1_283': {'h': 1, 'w': 0}, # 冰霜元素 655
  'EX1_284': {'h': 0, 'w': 0}, # 蒼藍龍 544
  'EX1_287': {'h': 0, 'w': 0}, # 法術反制 3XX
  'EX1_294': {'h': 1, 'w': 0}, # 鏡像體 3XX
  'EX1_295': {'h': 0, 'w': 0}, # 寒冰屏障 3XX
  'EX1_298': {'h': 0, 'w': 0}, # 『炎魔』拉格納羅斯 888
  'EX1_301': {'h': 0, 'w': 0}, # 惡魔守衛 335
  'EX1_303': {'h': 1, 'w': 0}, # 暗影之焰 4XX
  'EX1_304': {'h': 1, 'w': 0}, # 虛無恐獸 333
  'EX1_309': {'h': 0, 'w': 0}, # 靈魂虹吸 6XX
  'EX1_310': {'h': 1, 'w': 0}, # 末日守衛 557
  'EX1_312': {'h': 0, 'w': 0}, # 扭曲虛空 8XX
  'EX1_313': {'h': 0, 'w': 0}, # 深淵領主 456
  'EX1_315': {'h': 1, 'w': 0}, # 召喚傳送門 404
  'EX1_316': {'h': 0, 'w': 0}, # 壓倒性的力量 1XX
  'EX1_319': {'h': 1, 'w': 0}, # 烈焰小鬼 132
  'EX1_320': {'h': 0, 'w': 0}, # 末日災厄 5XX
  'EX1_323': {'h': 0, 'w': 0}, # 賈拉克瑟斯領主 9315
  'EX1_332': {'h': 0, 'w': 0}, # 沉默 0XX
  'EX1_334': {'h': 0, 'w': 0}, # 暗影狂亂 4XX
  'EX1_341': {'h': 1, 'w': 0}, # 光束泉 205
  'EX1_345': {'h': 0, 'w': 0}, # 心理遊戲 4XX
  'EX1_349': {'h': 0, 'w': 0}, # 神恩術 3XX
  'EX1_350': {'h': 0, 'w': 0}, # 預言者費倫 777
  'EX1_354': {'h': 0, 'w': 0}, # 聖療術 8XX
  'EX1_355': {'h': 1, 'w': 0}, # 勇者祝福 5XX
  'EX1_362': {'h': 0, 'w': 0}, # 銀色黎明保衛者 222
  'EX1_363': {'h': 0, 'w': 0}, # 智慧祝福 1XX
  'EX1_365': {'h': 1, 'w': 0}, # 神聖憤怒 5XX
  'EX1_366': {'h': 0, 'w': 0}, # 正義之劍 315
  'EX1_379': {'h': 1, 'w': 0}, # 懺悔 1XX
  'EX1_382': {'h': 1, 'w': 0}, # 奧多爾保安官 333
  'EX1_384': {'h': 0, 'w': 0}, # 復仇之怒 6XX
  'EX1_391': {'h': 0, 'w': 0}, # 猛擊 2XX
  'EX1_392': {'h': 0, 'w': 0}, # 戰鬥狂怒 2XX
  'EX1_393': {'h': 0, 'w': 0}, # 阿曼尼狂戰士 223
  'EX1_402': {'h': 0, 'w': 0}, # 護甲鍛造師 214
  'EX1_405': {'h': 0, 'w': 0}, # 執盾兵 104
  'EX1_407': {'h': 0, 'w': 0}, # 鬥毆 5XX
  'EX1_408': {'h': 0, 'w': 0}, # 致死打擊 4XX
  'EX1_409': {'h': 0, 'w': 0}, # 升級！ 1XX
  'EX1_410': {'h': 0, 'w': 0}, # 盾牌猛擊 1XX
  'EX1_411': {'h': 0, 'w': 0}, # 血吼之斧 771
  'EX1_414': {'h': 0, 'w': 0}, # 葛羅瑪許·地獄吼 849
  'EX1_507': {'h': 0, 'w': 0}, # 魚人隊長 333
  'EX1_522': {'h': 1, 'w': 0}, # 有耐心的刺客 211
  'EX1_531': {'h': 1, 'w': 0}, # 食腐土狼 222
  'EX1_533': {'h': 0, 'w': 0}, # 誤導 2XX
  'EX1_534': {'h': 0, 'w': 0}, # 長鬃草原獅 665
  'EX1_536': {'h': 0, 'w': 0}, # 鷹角弓 332
  'EX1_537': {'h': 0, 'w': 0}, # 爆裂射擊 5XX
  'EX1_543': {'h': 0, 'w': 0}, # 克洛許王 988
  'EX1_544': {'h': 0, 'w': 0}, # 照明彈 2XX
  'EX1_549': {'h': 0, 'w': 0}, # 狂野怒火 1XX
  'EX1_554': {'h': 0, 'w': 0}, # 毒蛇陷阱 2XX
  'EX1_557': {'h': 0, 'w': 0}, # 納特·帕格 204
  'EX1_558': {'h': 0, 'w': 0}, # 哈里遜·瓊斯 554
  'EX1_559': {'h': 0, 'w': 0}, # 大法師安東尼達斯 757
  'EX1_560': {'h': 0, 'w': 0}, # 諾茲多姆 988
  'EX1_561': {'h': 0, 'w': 0}, # 雅立史卓莎 988
  'EX1_562': {'h': 0, 'w': 0}, # 奧妮克希亞 988
  'EX1_563': {'h': 0, 'w': 0}, # 瑪里苟斯 9412
  'EX1_564': {'h': 0, 'w': 0}, # 無面操縱者 533
  'EX1_567': {'h': 1, 'w': 0}, # 末日錘 528
  'EX1_570': {'h': 0, 'w': 0}, # 撕咬 4XX
  'EX1_571': {'h': 1, 'w': 0}, # 自然之力 6XX
  'EX1_572': {'h': 0, 'w': 0}, # 伊瑟拉 9412
  'EX1_573': {'h': 0, 'w': 0}, # 塞納留斯 958
  'EX1_575': {'h': 0, 'w': 0}, # 法力之潮圖騰 303
  'EX1_577': {'h': 0, 'w': 0}, # 比斯巨獸 697
  'EX1_578': {'h': 0, 'w': 0}, # 兇蠻 1XX
  'EX1_584': {'h': 0, 'w': 0}, # 老邁的法師 425
  'EX1_586': {'h': 0, 'w': 0}, # 海巨人 1088
  'EX1_590': {'h': 0, 'w': 0}, # 血騎士 333
  'EX1_591': {'h': 1, 'w': 0}, # 奧奇奈靈魂牧師 435
  'EX1_594': {'h': 1, 'w': 0}, # 氣化 3XX
  'EX1_595': {'h': 0, 'w': 0}, # 教派宗師 442
  'EX1_596': {'h': 0, 'w': 0}, # 惡魔火焰 2XX
  'EX1_597': {'h': 0, 'w': 0}, # 小鬼召喚師 315
  'EX1_603': {'h': 0, 'w': 0}, # 殘酷的監工 222
  'EX1_609': {'h': 1, 'w': 0}, # 狙擊 2XX
  'EX1_610': {'h': 1, 'w': 0}, # 爆炸陷阱 2XX
  'EX1_611': {'h': 1, 'w': 0}, # 冰凍陷阱 2XX
  'EX1_612': {'h': 0, 'w': 0}, # 祈倫托法師 343
  'EX1_613': {'h': 0, 'w': 0}, # 艾德溫·范克里夫 322
  'EX1_614': {'h': 0, 'w': 0}, # 伊利丹·怒風 675
  'EX1_616': {'h': 0, 'w': 0}, # 法力怨靈 222
  'EX1_617': {'h': 0, 'w': 0}, # 致命射擊 3XX
  'EX1_619': {'h': 1, 'w': 0}, # 一視同仁 2XX
  'EX1_620': {'h': 0, 'w': 0}, # 熔火巨人 2088
  'EX1_623': {'h': 1, 'w': 0}, # 神殿執行者 666
  'EX1_625': {'h': 0, 'w': 0}, # 暗影形態 3XX
  'EX1_626': {'h': 0, 'w': 0}, # 群體驅魔 4XX
  'NEW1_005': {'h': 1, 'w': 0}, # 綁匪 653
  'NEW1_007': {'h': 1, 'w': 0}, # 星殞術 5XX
  'NEW1_008': {'h': 1, 'w': 0}, # 知識古樹 755
  'NEW1_010': {'h': 0, 'w': 0}, # 『馭風者』奧拉基爾 835
  'NEW1_012': {'h': 0, 'w': 0}, # 法力龍鰻 113
  'NEW1_014': {'h': 1, 'w': 0}, # 偽裝大師 444
  'NEW1_017': {'h': 0, 'w': 0}, # 飢餓的螃蟹 112
  'NEW1_018': {'h': 0, 'w': 0}, # 血帆劫掠者 223
  'NEW1_019': {'h': 1, 'w': 0}, # 飛刀手 232
  'NEW1_020': {'h': 1, 'w': 0}, # 狂野火占師 232
  'NEW1_021': {'h': 0, 'w': 0}, # 末日預言者 207
  'NEW1_024': {'h': 0, 'w': 0}, # 綠皮船長 554
  'NEW1_025': {'h': 1, 'w': 0}, # 血帆海寇 112
  'NEW1_026': {'h': 0, 'w': 0}, # 紫羅蘭教師 435
  'NEW1_027': {'h': 0, 'w': 0}, # 南海船長 333
  'NEW1_029': {'h': 0, 'w': 0}, # 米歐浩斯·曼納斯頓 244
  'NEW1_030': {'h': 0, 'w': 0}, # 死亡之翼 101212
  'NEW1_036': {'h': 0, 'w': 0}, # 命令之吼 2XX
  'NEW1_037': {'h': 0, 'w': 0}, # 劍類鍛造大師 213
  'NEW1_038': {'h': 0, 'w': 0}, # 戈魯爾 877
  'NEW1_040': {'h': 0, 'w': 0}, # 霍格 644
  'NEW1_041': {'h': 0, 'w': 0}, # 奔竄的科多獸 535
  'tt_004': {'h': 1, 'w': 0}, # 食肉食屍鬼 323
  'tt_010': {'h': 0, 'w': 0}, # 法術扭曲者 3XX
  # GVG
  'GVG_001': {'h': 0, 'w': 0}, # 烈焰火砲 2XX
  'GVG_002': {'h': 0, 'w': 0}, # 剷雪機器人 223
  'GVG_003': {'h': 0, 'w': 0}, # 不穩定的傳送門 2XX
  'GVG_004': {'h': 0, 'w': 0}, # 哥布林轟炸法師 454
  'GVG_005': {'h': 0, 'w': 0}, # 麥迪文的回音 4XX
  'GVG_006': {'h': 0, 'w': 0}, # 機械召喚師 223
  'GVG_007': {'h': 0, 'w': 0}, # 烈焰戰輪 777
  'GVG_008': {'h': 0, 'w': 0}, # 聖光炸彈 6XX
  'GVG_009': {'h': 0, 'w': 0}, # 暗影炸彈手 121
  'GVG_010': {'h': 0, 'w': 0}, # 費倫的祝福 3XX
  'GVG_011': {'h': 0, 'w': 0}, # 縮小射線工程師 232
  'GVG_012': {'h': 0, 'w': 0}, # 那魯之光 1XX
  'GVG_013': {'h': 0, 'w': 0}, # 齒輪大師 112
  'GVG_014': {'h': 0, 'w': 0}, # 沃金 562
  'GVG_015': {'h': 0, 'w': 0}, # 黑暗炸彈 2XX
  'GVG_016': {'h': 0, 'w': 0}, # 惡魔劫奪者 588
  'GVG_017': {'h': 0, 'w': 0}, # 召喚寵物 2XX
  'GVG_018': {'h': 0, 'w': 0}, # 苦痛仕女 214
  'GVG_019': {'h': 0, 'w': 0}, # 惡魔之心 5XX
  'GVG_020': {'h': 0, 'w': 0}, # 惡魔火砲 435
  'GVG_021': {'h': 0, 'w': 0}, # 瑪爾加尼斯 997
  'GVG_022': {'h': 0, 'w': 0}, # 技工的磨刀油 4XX
  'GVG_023': {'h': 0, 'w': 0}, # 哥布林自動理髮師 232
  'GVG_024': {'h': 0, 'w': 0}, # 齒輪大師的扳手 313
  'GVG_025': {'h': 0, 'w': 0}, # 獨眼騙子 241
  'GVG_026': {'h': 0, 'w': 0}, # 假死 2XX
  'GVG_027': {'h': 0, 'w': 0}, # 鋼鐵師尊 322
  'GVG_028': {'h': 0, 'w': 0}, # 貿易親王加里維克斯 658
  'GVG_029': {'h': 0, 'w': 0}, # 先祖之喚 4XX
  'GVG_030': {'h': 1, 'w': 0}, # 電鍍機械小熊 222
  'GVG_031': {'h': 0, 'w': 0}, # 回收 6XX
  'GVG_032': {'h': 1, 'w': 0}, # 林地看管者 324
  'GVG_033': {'h': 0, 'w': 0}, # 生命之樹 9XX
  'GVG_034': {'h': 0, 'w': 0}, # 機械熊-貓 676
  'GVG_035': {'h': 0, 'w': 0}, # 瑪洛尼 797
  'GVG_036': {'h': 0, 'w': 0}, # 動力戰錘 332
  'GVG_037': {'h': 0, 'w': 0}, # 漩渦打擊裝置 232
  'GVG_038': {'h': 0, 'w': 0}, # 轟雷 2XX
  'GVG_039': {'h': 0, 'w': 0}, # 活力圖騰 203
  'GVG_040': {'h': 0, 'w': 0}, # 沙鰭靈行者 425
  'GVG_041': {'h': 0, 'w': 0}, # 黑暗幽光 6XX
  'GVG_042': {'h': 0, 'w': 0}, # 奈普圖隆 777
  'GVG_043': {'h': 0, 'w': 0}, # 旋刃火箭筒 222
  'GVG_044': {'h': 0, 'w': 0}, # 蜘蛛坦克 334
  'GVG_045': {'h': 0, 'w': 0}, # 小鬼爆破 4XX
  'GVG_046': {'h': 0, 'w': 0}, # 萬獸之王 526
  'GVG_047': {'h': 0, 'w': 0}, # 破壞工作 4XX
  'GVG_048': {'h': 1, 'w': 0}, # 鋼牙獸 333
  'GVG_049': {'h': 0, 'w': 0}, # 加茲瑞拉 769
  'GVG_050': {'h': 0, 'w': 0}, # 彈跳鋒刃 3XX
  'GVG_051': {'h': 0, 'w': 0}, # 戰爭機器人 113
  'GVG_052': {'h': 0, 'w': 0}, # 粉碎 7XX
  'GVG_053': {'h': 0, 'w': 0}, # 女盾侍 655
  'GVG_054': {'h': 0, 'w': 0}, # 巨魔戰槌 342
  'GVG_055': {'h': 0, 'w': 0}, # 破舊的維修機甲 425
  'GVG_056': {'h': 0, 'w': 0}, # 鋼鐵破滅邪神 665
  'GVG_057': {'h': 0, 'w': 0}, # 光明聖印 2XX
  'GVG_058': {'h': 0, 'w': 0}, # 護盾小機器人 222
  'GVG_059': {'h': 0, 'w': 0}, # 齒輪巨錘 323
  'GVG_060': {'h': 0, 'w': 0}, # 軍需官 525
  'GVG_061': {'h': 1, 'w': 0}, # 整裝備戰 3XX
  'GVG_062': {'h': 0, 'w': 0}, # 鈷藍守護者 563
  'GVG_063': {'h': 0, 'w': 0}, # 伯瓦爾·弗塔根 517
  'GVG_064': {'h': 0, 'w': 0}, # 淤泥踐踏者 232
  'GVG_065': {'h': 0, 'w': 0}, # 巨魔蠻卒 344
  'GVG_066': {'h': 0, 'w': 0}, # 砂槌薩滿 454
  'GVG_067': {'h': 0, 'w': 0}, # 石裂穴居怪 223
  'GVG_068': {'h': 0, 'w': 0}, # 魁梧的石顎穴居怪 435
  'GVG_069': {'h': 0, 'w': 0}, # 古董治療機器人 533
  'GVG_070': {'h': 0, 'w': 0}, # 老水手 574
  'GVG_071': {'h': 0, 'w': 0}, # 迷路的陸行鳥 454
  'GVG_072': {'h': 0, 'w': 0}, # 暗影拳擊手 223
  'GVG_073': {'h': 0, 'w': 0}, # 眼鏡蛇射擊 5XX
  'GVG_074': {'h': 0, 'w': 0}, # 凱贊秘術使 443
  'GVG_075': {'h': 0, 'w': 0}, # 船艦主砲 223
  'GVG_077': {'h': 0, 'w': 0}, # 憎惡魔像 699
  'GVG_078': {'h': 0, 'w': 0}, # 機械雪人 445
  'GVG_079': {'h': 0, 'w': 0}, # 超能坦克麥克斯 877
  'GVG_080': {'h': 0, 'w': 0}, # 尖牙德魯伊 544
  'GVG_081': {'h': 0, 'w': 0}, # 吉爾布林潛獵者 223
  'GVG_082': {'h': 0, 'w': 0}, # 發條地精 121
  'GVG_083': {'h': 0, 'w': 0}, # 升級版修理機器人 555
  'GVG_084': {'h': 0, 'w': 0}, # 飛行器 314
  'GVG_085': {'h': 0, 'w': 0}, # 煩人機器人 212
  'GVG_086': {'h': 0, 'w': 0}, # 攻城機具 555
  'GVG_087': {'h': 0, 'w': 0}, # 熱砂狙擊手 223
  'GVG_088': {'h': 0, 'w': 0}, # 巨魔忍者 566
  'GVG_089': {'h': 0, 'w': 0}, # 光明引路者 324
  'GVG_090': {'h': 0, 'w': 0}, # 更瘋狂的炸彈客 554
  'GVG_091': {'h': 0, 'w': 0}, # 秘法剋星X-21 425
  'GVG_092': {'h': 0, 'w': 0}, # 地精實驗家 332
  'GVG_093': {'h': 0, 'w': 0}, # 訓練假人 002
  'GVG_094': {'h': 0, 'w': 0}, # 吉福斯 414
  'GVG_095': {'h': 0, 'w': 0}, # 哥布林工兵 324
  'GVG_096': {'h': 1, 'w': 0}, # 有駕駛的伐木機 443
  'GVG_097': {'h': 0, 'w': 0}, # 小小驅魔者 323
  'GVG_098': {'h': 0, 'w': 0}, # 諾姆瑞根步兵 314
  'GVG_099': {'h': 0, 'w': 0}, # 投彈手 533
  'GVG_101': {'h': 0, 'w': 0}, # 血色淨化者 343
  'GVG_102': {'h': 1, 'w': 0}, # 地精區技師 333
  'GVG_103': {'h': 0, 'w': 0}, # 微型裝甲 212
  'GVG_104': {'h': 0, 'w': 0}, # 大哥布林 323
  'GVG_105': {'h': 0, 'w': 0}, # 有駕駛的飛天魔像 664
  'GVG_106': {'h': 0, 'w': 0}, # 拾荒機器人 515
  'GVG_107': {'h': 0, 'w': 0}, # 強化機器人 432
  'GVG_108': {'h': 0, 'w': 0}, # 重組轉化師 232
  'GVG_109': {'h': 0, 'w': 0}, # 迷你法師 441
  'GVG_110': {'h': 0, 'w': 0}, # 爆爆博士 777
  'GVG_111': {'h': 0, 'w': 0}, # 彌米倫之首 545
  'GVG_112': {'h': 0, 'w': 0}, # 巨魔莫古 676
  'GVG_113': {'h': 0, 'w': 0}, # 敵人收割者4000 869
  'GVG_114': {'h': 0, 'w': 0}, # 斯尼德的伐木機 857
  'GVG_115': {'h': 0, 'w': 0}, # 托斯利 657
  'GVG_116': {'h': 0, 'w': 0}, # 機電師瑟瑪普拉格 997
  'GVG_117': {'h': 0, 'w': 0}, # 加茲魯維 636
  'GVG_118': {'h': 0, 'w': 0}, # 特洛格佐爾 766
  'GVG_119': {'h': 0, 'w': 0}, # 布靈登3000型 534
  'GVG_120': {'h': 0, 'w': 0}, # 赫米特·奈辛瓦里 563
  'GVG_121': {'h': 0, 'w': 0}, # 發條巨人 1288
  'GVG_122': {'h': 0, 'w': 0}, # 嬌小的法術干擾師 425
  'GVG_123': {'h': 0, 'w': 0}, # 煤煙噴吐機器人 333
  # TGT
  'AT_001': {'h': 0, 'w': 0}, # 火焰長矛 5XX
  'AT_002': {'h': 0, 'w': 0}, # 火焰稻草人 3XX
  'AT_003': {'h': 0, 'w': 0}, # 陣亡英雄之靈 232
  'AT_004': {'h': 0, 'w': 0}, # 秘法衝擊 1XX
  'AT_005': {'h': 0, 'w': 0}, # 變形術：野豬 3XX
  'AT_006': {'h': 1, 'w': 0}, # 達拉然志士 435
  'AT_007': {'h': 0, 'w': 0}, # 魔法鏢客 334
  'AT_008': {'h': 0, 'w': 0}, # 凜懼島飛龍 666
  'AT_009': {'h': 0, 'w': 0}, # 羅甯 877
  'AT_010': {'h': 0, 'w': 0}, # 山羊牧人 533
  'AT_011': {'h': 0, 'w': 0}, # 神聖勇士 435
  'AT_012': {'h': 0, 'w': 0}, # 暗影爪牙 454
  'AT_013': {'h': 0, 'w': 0}, # 真言術：耀 1XX
  'AT_014': {'h': 0, 'w': 0}, # 暗影惡魔 333
  'AT_015': {'h': 0, 'w': 0}, # 歸順 2XX
  'AT_016': {'h': 0, 'w': 0}, # 混亂 2XX
  'AT_017': {'h': 0, 'w': 0}, # 暮光守護者 426
  'AT_018': {'h': 0, 'w': 0}, # 告解者帕爾璀絲 754
  'AT_019': {'h': 0, 'w': 0}, # 恐懼戰馬 411
  'AT_020': {'h': 0, 'w': 0}, # 可怕的末日守衛 768
  'AT_021': {'h': 0, 'w': 0}, # 小小邪惡騎士 232
  'AT_022': {'h': 0, 'w': 0}, # 賈拉克瑟斯之拳 4XX
  'AT_023': {'h': 0, 'w': 0}, # 虛無粉碎者 654
  'AT_024': {'h': 0, 'w': 0}, # 惡魔融合 2XX
  'AT_025': {'h': 0, 'w': 0}, # 黑暗交易 6XX
  'AT_026': {'h': 0, 'w': 0}, # 憤怒守衛 243
  'AT_027': {'h': 0, 'w': 0}, # 威爾弗雷德‧菲斯巴恩 644
  'AT_028': {'h': 0, 'w': 0}, # 影潘騎士 537
  'AT_029': {'h': 0, 'w': 0}, # 海賊 121
  'AT_030': {'h': 0, 'w': 0}, # 幽暗城驍士 232
  'AT_031': {'h': 0, 'w': 0}, # 扒手 222
  'AT_032': {'h': 0, 'w': 0}, # 黑市商人 343
  'AT_033': {'h': 0, 'w': 0}, # 盜竊 3XX
  'AT_034': {'h': 0, 'w': 0}, # 毒刃 413
  'AT_035': {'h': 0, 'w': 0}, # 地底潛伏 3XX
  'AT_036': {'h': 0, 'w': 0}, # 阿努巴拉克 984
  'AT_037': {'h': 0, 'w': 0}, # 糾纏之根 1XX
  'AT_038': {'h': 0, 'w': 0}, # 達納蘇斯志士 223
  'AT_039': {'h': 1, 'w': 0}, # 蠻荒戰鬥者 454
  'AT_041': {'h': 0, 'w': 0}, # 荒野騎士 766
  'AT_042': {'h': 0, 'w': 0}, # 刃牙德魯伊 221
  'AT_043': {'h': 0, 'w': 0}, # 星體共融 4XX
  'AT_044': {'h': 0, 'w': 0}, # 堆肥 3XX
  'AT_045': {'h': 0, 'w': 0}, # 艾維娜 955
  'AT_046': {'h': 0, 'w': 0}, # 巨牙圖騰師 332
  'AT_047': {'h': 0, 'w': 0}, # 德萊尼圖騰雕刻師 444
  'AT_048': {'h': 0, 'w': 0}, # 治療波 3XX
  'AT_049': {'h': 0, 'w': 0}, # 雷霆崖驍士 536
  'AT_050': {'h': 0, 'w': 0}, # 充能戰錘 424
  'AT_051': {'h': 0, 'w': 0}, # 元素毀滅 3XX
  'AT_052': {'h': 1, 'w': 0}, # 圖騰魔像 234
  'AT_053': {'h': 0, 'w': 0}, # 先祖知識 2XX
  'AT_054': {'h': 0, 'w': 0}, # 喚霧者 644
  'AT_055': {'h': 0, 'w': 0}, # 快速治療 1XX
  'AT_056': {'h': 0, 'w': 0}, # 強力射擊 3XX
  'AT_057': {'h': 0, 'w': 0}, # 獸欄管理員 342
  'AT_058': {'h': 0, 'w': 0}, # 國王的伊萊克 232
  'AT_059': {'h': 0, 'w': 0}, # 勇敢弓箭手 121
  'AT_060': {'h': 0, 'w': 0}, # 放熊陷阱 2XX
  'AT_061': {'h': 0, 'w': 0}, # 全面備戰 2XX
  'AT_062': {'h': 0, 'w': 0}, # 蜘蛛囊 6XX
  'AT_063': {'h': 0, 'w': 0}, # 酸喉 742
  'AT_063t': {'h': 0, 'w': 0}, # 懼鱗 342
  'AT_064': {'h': 0, 'w': 0}, # 重擊 3XX
  'AT_065': {'h': 0, 'w': 0}, # 王家防衛者 332
  'AT_066': {'h': 0, 'w': 0}, # 奧格瑪志士 333
  'AT_067': {'h': 0, 'w': 0}, # 猛瑪象人首領 453
  'AT_068': {'h': 0, 'w': 0}, # 提振士氣 2XX
  'AT_069': {'h': 0, 'w': 0}, # 練習夥伴 232
  'AT_070': {'h': 0, 'w': 0}, # 天空隊長克拉格 746
  'AT_071': {'h': 0, 'w': 0}, # 雅立史卓莎的勇士 223
  'AT_072': {'h': 0, 'w': 0}, # 瓦里安‧烏瑞恩 1077
  'AT_073': {'h': 0, 'w': 0}, # 運動精神 1XX
  'AT_074': {'h': 0, 'w': 0}, # 勇士徽印 3XX
  'AT_075': {'h': 0, 'w': 0}, # 戰馬訓練師 324
  'AT_076': {'h': 0, 'w': 0}, # 魚人騎士 434
  'AT_077': {'h': 0, 'w': 0}, # 銀白長槍 222
  'AT_078': {'h': 0, 'w': 0}, # 高手過招 6XX
  'AT_079': {'h': 0, 'w': 0}, # 神秘挑戰者 666
  'AT_080': {'h': 0, 'w': 0}, # 要塞指揮官 223
  'AT_081': {'h': 0, 'w': 0}, # 『純淨者』埃卓克 737
  'AT_082': {'h': 0, 'w': 0}, # 低階侍從 112
  'AT_083': {'h': 0, 'w': 0}, # 龍鷹騎士 333
  'AT_084': {'h': 0, 'w': 0}, # 槍僮 212
  'AT_085': {'h': 0, 'w': 0}, # 湖中少女 426
  'AT_086': {'h': 0, 'w': 0}, # 破壞者 343
  'AT_087': {'h': 0, 'w': 0}, # 銀白騎兵 321
  'AT_088': {'h': 0, 'w': 0}, # 莫古的勇士 685
  'AT_089': {'h': 0, 'w': 0}, # 骨衛中尉 232
  'AT_090': {'h': 0, 'w': 0}, # 穆克拉的勇士 543
  'AT_091': {'h': 0, 'w': 0}, # 聯賽醫護兵 418
  'AT_092': {'h': 1, 'w': 0}, # 寒冰狂怒者 352
  'AT_093': {'h': 0, 'w': 0}, # 嚴寒狗頭人 426
  'AT_094': {'h': 0, 'w': 0}, # 火焰雜耍師 223
  'AT_095': {'h': 0, 'w': 0}, # 靜默騎士 322
  'AT_096': {'h': 0, 'w': 0}, # 發條騎士 555
  'AT_097': {'h': 0, 'w': 0}, # 聯賽觀眾 121
  'AT_098': {'h': 0, 'w': 0}, # 雜耍吞法者 665
  'AT_099': {'h': 0, 'w': 0}, # 科多獸騎士 635
  'AT_100': {'h': 0, 'w': 0}, # 白銀之手長官 333
  'AT_101': {'h': 0, 'w': 0}, # 鬥技場鬥士 556
  'AT_102': {'h': 0, 'w': 0}, # 捕獲的蟄猛巨蟲 759
  'AT_103': {'h': 0, 'w': 0}, # 北海海怪 997
  'AT_104': {'h': 0, 'w': 0}, # 巨牙矛騎兵 555
  'AT_105': {'h': 0, 'w': 0}, # 受傷的科瓦迪爾 124
  'AT_106': {'h': 0, 'w': 0}, # 聖光勇士 343
  'AT_108': {'h': 0, 'w': 0}, # 裝甲戰馬 453
  'AT_109': {'h': 0, 'w': 0}, # 銀白巡邏兵 224
  'AT_110': {'h': 0, 'w': 0}, # 大競技場經理 325
  'AT_111': {'h': 0, 'w': 0}, # 餐點小販 435
  'AT_112': {'h': 0, 'w': 0}, # 至尊矛騎兵 656
  'AT_113': {'h': 0, 'w': 0}, # 募兵官 554
  'AT_114': {'h': 0, 'w': 0}, # 邪惡挑釁者 454
  'AT_115': {'h': 0, 'w': 0}, # 擊劍教練 322
  'AT_116': {'h': 0, 'w': 0}, # 龍眠使者 214
  'AT_117': {'h': 0, 'w': 0}, # 大會主持人 342
  'AT_118': {'h': 0, 'w': 0}, # 大十字軍 655
  'AT_119': {'h': 0, 'w': 0}, # 科瓦迪爾劫掠者 544
  'AT_120': {'h': 0, 'w': 0}, # 冰霜巨人 1088
  'AT_121': {'h': 0, 'w': 0}, # 大明星 444
  'AT_122': {'h': 0, 'w': 0}, # 『穿刺者』戈莫克 444
  'AT_123': {'h': 0, 'w': 0}, # 寒冽之喉 766
  'AT_124': {'h': 0, 'w': 0}, # 波爾夫‧拉姆榭 639
  'AT_125': {'h': 0, 'w': 0}, # 冰嚎 91010
  'AT_127': {'h': 0, 'w': 0}, # 奈薩斯勇士薩拉德 545
  'AT_128': {'h': 0, 'w': 0}, # 骷髏騎士 674
  'AT_129': {'h': 0, 'w': 0}, # 菲歐拉‧光寂 334
  'AT_130': {'h': 0, 'w': 0}, # 海劫者 667
  'AT_131': {'h': 0, 'w': 0}, # 艾狄絲‧暗寂 334
  'AT_132': {'h': 0, 'w': 0}, # 審判者瑪瑞爾 663
  'AT_133': {'h': 0, 'w': 0}, # 加基森矛騎兵 112
  # Reward + Promotion
  'EX1_062': {'h': 0, 'w': 0}, # 老瞎眼 424
  'NEW1_016': {'h': 0, 'w': 0}, # 船長的鸚鵡 211
  'EX1_112': {'h': 0, 'w': 0}, # 傑爾賓·梅卡托克 666
  'PRO_001': {'h': 0, 'w': 0}, # 精英牛頭大佬 555

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
  'Common': '普通',
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


def getEnameDict():
  foo = {}
  txtPath = "%s\Data\Win\cardxml0\CAB-cardxml0\TextAsset" % (hearthstonePath) 
  doc = minidom.parse("%s\enUS.txt" % (txtPath))
  root = doc.documentElement
  cards = root.getElementsByTagName("Entity")
  for card in cards:
    foo[card.getAttribute('CardID')] = getTagValue(card, "CardName")
  return foo

# parse *.txt to xlsx
def parseTxtFiles():
  edict = getEnameDict()

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
        or (re.match(".+_\d{3}t$", cid) and not cid in ['AT_063t']) \
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
    elif cset in []: # ['Curse of Naxxramas', 'Blackrock Mountain']: # TODO: no gold
      format = green
      dust = ''
      have = total
    elif cset in ['Curse of Naxxramas', 'Blackrock Mountain']: 
      format = red
      dust = ''
      have = 0
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
    worksheet.write(i, 29, edict.get(cid)) 

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
  parseTxtFiles()

# card dust
# http://tinyurl.com/ovmmvrk
