#!/usr/bin/python3
import os
import openpyxl
import config
import dv
import ebird
import json

x = ebird.ebird(config.token)

workbook = openpyxl.load_workbook('/Users/taqini/Documents/植物园鸟调表格/temp.xlsx')

# generate by kimi
index = {"月季园": ["D", "E", "F"],
         "温室西": ["G", "H", "I"],
         "芍药园-牡丹园-海棠园": ["J", "K", "L"],
         "梅园-栈道": ["M", "N", "O"],
         "水源头": ["P", "Q", "R"],
         "碧桃园-丁香园": ["S", "T", "U"],
         "卧佛寺周边": ["V", "W", "X"],
         "绚秋苑": ["Y", "Z", "AA"],
         "东南门+澄静湖": ["AB", "AC", "AD"],
         "黄叶村+中湖": ["AE", "AF", "AG"],
         "王锡彤墓+北湖": ["AH", "AI", "AJ"],
         "梁启超墓": ["AK", "AL", "AM"],
         "树木园": ["AN", "AO", "AP"]
         }

result = {}

sheet = workbook.active
sheet.title = config.date
print(sheet)

# 基础信息
sheet['C2'] = config.date
sheet['C3'] = config.volunters
sheet['C4'] = config.howManyPeople
sheet['F2'] = config.startTime
sheet['F3'] = config.endTime
sheet['K4'] = config.weather

# 初始行号
birdIndex = 1
birdList = {}

def formSum(idx):
	s = "="
	for i in index:
		s += index[i][0]+str(idx)+'+'
	return s[:-1]

# 获取checklists中所有鸟种，并编号id
for cList in config.checklist:
	cId = config.checklist[cList]
	if cId:
		print("[*] 获取%s的checklist=%s中..."%(cList,cId))
		detail = x.get_report_detail(subId=cId)
		res = x.getCount(detail)
		result[cList] = res
		# print(cList, cId, index[cList])
		for i in res:
			comName, howManyStr = i[0],i[1]
			if comName not in birdList.keys():
				birdList[comName] = birdIndex
				sheet['A%d'%(birdIndex+6)] = birdIndex
				sheet['B%d'%(birdIndex+6)] = comName
				sheet['C%d'%(birdIndex+6)] = formSum(birdIndex+6)
				birdIndex += 1

			# print(comName, howManyStr)
		# print(birdList)

print("[+] 已获取%d个checklist，总鸟种数%d"%(len(result),birdIndex-1))
# print(result)

def queryCountByComName(res, comName):
	for i in res:
		if i[0] == comName:
			return int(i[1])
	else:
		return 0


# 逐行写入鸟种信息
print("[*] 鸟种信息写入中...")
# print(birdList)

print('[#] 自动填充鸟种行为及生境中...')

for bird in birdList:
	bId = birdList[bird]
	# print(bird, bId)
	for r in result:
		count = queryCountByComName(result[r], bird)
		if count != 0:
			# print(r, bird, count)
			sheet['%s%d'%(index[r][0],bId+6)] = count
			be = sheet['%s%d'%(index[r][1],bId+6)]
			li = sheet['%s%d'%(index[r][2],bId+6)]
			if bird in dv.dBeLi:
				be.value = dv.dBeLi[bird][0]
				li.value = dv.dBeLi[bird][1]
				if bird in dv.ListWater:
					dv.addDV(dv.dvbe1, dv.dvli1, be, li)
				else:
					dv.addDV(dv.dvbe2, dv.dvli2, be, li)
				print("[%s]-<%s,%s>"%(bird,be.value,li.value),end=",")
			else:
				print("[!] 请在dv中补充[%s]行为及生境"%bird)
			# be = ""
			# li = ""

print("\n[+] 写入完毕！")

# 鸟种数
sheet['F4'] = birdIndex-1
sheet['I4'] = "=SUM(C7:C%d)"%(birdIndex+5)

# 水鸟行为、生境约束
sheet.add_data_validation(dv.dvbe1)
sheet.add_data_validation(dv.dvli1)
# 林鸟行为、生境约束
sheet.add_data_validation(dv.dvbe2)
sheet.add_data_validation(dv.dvli2)

# 保存
workbook.save('植物园鸟调-%s.xlsx'%config.date)