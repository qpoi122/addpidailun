# -*- coding: UTF-8 -*-

'''
__author__="zf"
__mtime__ = '2016/11/8/21/38'
__des__: 简单的读取文件
__lastchange__:'2016/11/16'
'''
from __future__ import division
import xlrd
import os
import math
from xlwt import Workbook, Formula
import xlrd
import sys
import types
import copy  


def is_chinese(uchar): 
        """判断一个unicode是否是汉字"""
        if uchar >= u'/u4e00' and uchar<=u'/u9fa5':
                return True
        else:
                return False

                
def is_num(unum):
	try:
		unum+1
	except TypeError:
		return 0
	else:
		return 1

#不带颜色的读取
def filename(content):
	#打开文件
	global workbook,file_excel
	file_excel=str(content)
	file=(file_excel+'.xls').decode('utf-8')#文件名及中文合理性
	if not os.path.exists(file):#判断文件是否存在
		file=(file_excel+'.xlsx').decode('utf-8')
		if not os.path.exists(file):
			print "文件不存在"
	workbook = xlrd.open_workbook(file)
	print 'suicce'


def readexcel(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："

	for name in range(len(Sheetname)):
		allneed=[]
		table = workbook.sheets()[name]
		nrows=table.nrows
		mid=[]
		for n in range(nrows):
		
			
		#获取每行内容
			a=table.row_values(n)
			small=[]
			print a,'aaaaaaaaaaaaaaaaaaaaaa'
			for i in range(len(a)):

				if a[1]==u'':
					print a,'aaaaaaaaaaa222222222222aaaaaaaaaaa'
					if mid!=[]:
						bom.append(mid)
						# 总数组里加上产品信息
					mid=[]
					c=[]
					c.append(a[0])				
					mid.append(c)
					break
				else:
				#获取单行内每一列的内容		
					if is_chinese(a[i]):
						a[i].encode('utf-8' )
					elif is_num(a[i])==1:
						if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
							a[i]=int(a[i])#将浮点数化成整数
					small.append(a[i])
					#获取每个零件的信息
				print small
			if n==nrows-1:
				bom.append(mid)
				print 'dasdasdasd'
			if small!=[]:
				mid.append(small)

			# 将所有的同一产品的零件放在一起
			

	print bom,'bbbbbbbbbbbbbbbbbbbbbb'



	for x in range(len(bom)):
		flag=[0,0,0]
		for y in range(len(bom[x])):
			if bom[x][y][0] in [u'ZP-2315C',u'ZP-382B',u'ZP-2315A']:
				bom[x][y][1] =u'打弯片'
			if len(bom[x][y])>1:
				if bom[x][y][1] ==u'皮带轮毂':
					flag[0]=1
				elif bom[x][y][1] ==u'打弯片':
					flag[1]=1
				elif bom[x][y][1] ==u'保护盖' and bom[x][y][0]==u'ZP-4509':
					flag[2]=1
		if flag==[0,0,0]:
			newbom.append(bom[x])
		elif flag==[1,0,0]:
			zf1=copy.deepcopy(bom[x])
			# print bom[x],'bomx'
			# print id(zf1),id(bom[x]),'idddddddddddddd'
			mid=zf1[1][0]+u'-黑'
			zf1[1][0]=mid
			zf1[0][0]=zf1[0][0]+u'-黑'
			print bom[x],'bomx'
			newbom.append(zf1)
			zf2=copy.deepcopy(bom[x])  
			# print id(zf2),id(bom[x]),'idddddddddddddd222'
			c2=zf2[1][0]+u'-白'
			zf2[1][0]=c2
			zf2[0][0]=zf2[0][0]+u'-白'
			newbom.append(zf2)
		elif flag==[1,1,0]:
			print bom[x],'bom1'
			zf1=copy.deepcopy(bom[x])
			zf1[1][0]=zf1[1][0]+u'-黑'
			for y in range(len(zf1)):
				if len(zf1[y])>1:
					if zf1[y][1]==u'打弯片':
						zf1[y][0]=zf1[y][0]+u'-白'
						break
			try:
				zf1[0][0]=zf1[0][0]+u'-黑白'
			except:
				zf1[0][0]=str(int(zf1[0][0]))+u'-黑白'
			newbom.append(zf1)
			print bom[x],'bom2'

			zf2=copy.deepcopy(bom[x])
			print bom[x],'bom3'
			zf2[1][0]=zf2[1][0]+u'-白'
			for y in range(len(zf2)):
				if len(zf2[y])>1:
					if zf2[y][1]==u'打弯片':
						zf2[y][0]=zf2[y][0]+u'-黑'
						break
			try:
				zf2[0][0]=zf2[0][0]+u'-白黑'
			except:
				zf2[0][0]=str(int(zf2[0][0]))+u'-白黑'
			newbom.append(zf2)
		elif flag==[1,0,1]:
			zf1=copy.deepcopy(bom[x])
			zf1[1][0]=zf1[1][0]+u'-黑'
			for y in range(len(zf1)):
				if len(zf1[y])>1:
					if u'保护盖' in zf1[y][1]:
						zf1[y][0]=zf1[y][0]+u'-无字'
						break
			try:
				zf1[0][0]=zf1[0][0]+u'-黑无'
			except:
				zf1[0][0]=str(int(zf1[0][0]))+u'-黑无'
			newbom.append(zf1)
			print bom[x],'bom2'

			zf2=copy.deepcopy(bom[x])
			print bom[x],'bom3'
			zf2[1][0]=zf2[1][0]+u'-白'
			for y in range(len(zf2)):
				if len(zf2[y])>1:
					if u'保护盖' in zf2[y][1]:
						zf2[y][0]=zf2[y][0]+u'-无字'
						break
			try:
				zf2[0][0]=zf2[0][0]+u'-白无'
			except:
				zf2[0][0]=str(int(zf2[0][0]))+u'-白无'
			newbom.append(zf2)

			zf3=copy.deepcopy(bom[x])
			zf3[1][0]=zf3[1][0]+u'-黑'
			for y in range(len(zf3)):
				if len(zf3[y])>1:
					if u'保护盖' in zf3[y][1]:
						zf3[y][0]=zf3[y][0]+u'-有字'
						break
			try:
				zf3[0][0]=zf3[0][0]+u'-黑有'
			except:
				zf3[0][0]=str(int(zf3[0][0]))+u'-黑有'
			newbom.append(zf3)
			print bom[x],'bom2'

			zf4=copy.deepcopy(bom[x])
			print bom[x],'bom3'
			zf4[1][0]=zf4[1][0]+u'-白'
			for y in range(len(zf4)):
				if len(zf4[y])>1:
					if u'保护盖' in zf4[y][1]:
						zf4[y][0]=zf4[y][0]+u'-有字'
						break
			try:
				zf4[0][0]=zf4[0][0]+u'-白有'
			except:
				zf4[0][0]=str(int(zf4[0][0]))+u'-白有'
			newbom.append(zf4)
		elif flag==[1,1,1]:
			zf1=copy.deepcopy(bom[x])
			zf1[1][0]=zf1[1][0]+u'-黑'
			for y in range(len(zf1)):
				if len(zf1[y])>1:
					if zf1[y][1]==u'打弯片':
						zf1[y][0]=zf1[y][0]+u'-白'
					if u'保护盖' in zf1[y][1]:
						zf1[y][0]=zf1[y][0]+u'-无字'
						break
			try:
				zf1[0][0]=zf1[0][0]+u'-黑白无'
			except:
				zf1[0][0]=str(int(zf1[0][0]))+u'-黑白无'
			newbom.append(zf1)
			print bom[x],'bom2'

			zf2=copy.deepcopy(bom[x])
			zf2[1][0]=zf2[1][0]+u'-黑'
			for y in range(len(zf2)):
				if len(zf2[y])>1:
					if zf2[y][1]==u'打弯片':
						zf2[y][0]=zf2[y][0]+u'-白'
					if u'保护盖' in zf2[y][1]:
						zf2[y][0]=zf2[y][0]+u'-有字'
						break
			try:
				zf2[0][0]=zf2[0][0]+u'-黑白有'
			except:
				zf2[0][0]=str(int(zf2[0][0]))+u'-黑白有'
			newbom.append(zf2)
			print bom[x],'bom2'

			zf3=copy.deepcopy(bom[x])
			zf3[1][0]=zf3[1][0]+u'-白'
			for y in range(len(zf3)):
				if len(zf3[y])>1:
					if zf3[y][1]==u'打弯片':
						zf3[y][0]=zf3[y][0]+u'-黑'
					if u'保护盖' in zf3[y][1]:
						zf3[y][0]=zf3[y][0]+u'-无字'
						break
			try:
				zf3[0][0]=zf3[0][0]+u'-白黑无'
			except:
				zf3[0][0]=str(int(zf3[0][0]))+u'-白黑无'
			newbom.append(zf3)
			print bom[x],'bom2'

			zf4=copy.deepcopy(bom[x])
			zf4[1][0]=zf4[1][0]+u'-白'
			for y in range(len(zf4)):
				if len(zf4[y])>1:
					if zf4[y][1]==u'打弯片':
						zf4[y][0]=zf4[y][0]+u'-黑'
					if u'保护盖' in zf4[y][1]:
						zf4[y][0]=zf4[y][0]+u'-有字'
						break
			try:
				zf4[0][0]=zf4[0][0]+u'-白黑有'
			except:
				zf4[0][0]=str(int(zf4[0][0]))+u'-白黑有'
			newbom.append(zf4)
			print bom[x],'bom2'


	out()
	fu()
	zi()











def zi():
	for x in range(len(newbom)):
		for y in range(len(newbom[x])):
			xinghao='ZNP-'+newbom[x][0][0]
			if len(newbom[x][y])>1:
				if u'工具'in newbom[x][1][1]:
					typename=u'单向皮带轮工具'
				else:
					typename=u'单向皮带轮总成'
				midzi=[]
				midzi.append(xinghao)
				midzi.append(typename)
				midzi.append(newbom[x][y][0])
				midzi.append(newbom[x][y][1])
				midzi.append(newbom[x][y][2])
				midzi.append('')
				midzi.append(newbom[x][y][6])
				midzi.append(newbom[x][y][7])
				zix.append(midzi)

	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	sheet1.write(0,0,u'子项模板')
	sheet1.write(1,0,u'父项存货编码')
	sheet1.write(1,1,u'父项存货全名')
	sheet1.write(1,2,u'存货编码')
	sheet1.write(1,3,u'存货全名')
	sheet1.write(1,4,u'数量')
	sheet1.write(1,5,u'损耗率')
	sheet1.write(1,6,u'备注1')
	sheet1.write(1,7,u'备注2')
	for i in range(len(zix)):
		for j in range (len(zix[i])):
			if is_chinese(zix[i][j]):
				zix[i][j].encode('utf-8')
			# elif not zix[i] and zix[i]!=0: 
			# 	print "空值",
			elif is_num(zix[i][j])==1:
				if math.modf(zix[i][j])[0]==0 or zix[i][j]==0:#获取数字的整数和小数
					zix[i][j]=int(zix[i][j])#将浮点数化成整数
			sheet1.write(i+2,j,zix[i][j])
	book.save('zibom_____3.xls')#存储excel
	book = xlrd.open_workbook('zibom_____3.xls')








				

def fu():

	filename('finalinput_cun')
		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："

	for name in range(len(Sheetname)):
		cun=[]
		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]	
		#获取单行内容
			a=table.row_values(n)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				
				mid.append(a[i])
			cun.append(mid)
	print cun,'bbbbbbbbbbbbbbbbbbbbbb'
	print newbom,'newbomnewbom'
	



	for x in range(len(newbom)):
		if u'工具' in newbom[x][1][1]:
			mid=[]
			bomtitile=u'ZNP-'+newbom[x][0][0]
			mid.append(bomtitile)	
			mid.append(bomtitile)
			mid.append('')
			mid.append(unicode('单向皮带轮工具BOM','utf-8'))
			mid.append(bomtitile)
			mid.append(unicode('单向皮带轮工具','utf-8'))
			mid.append(u'1')
			bombian=bombian+1
			fux.append(mid)

			#对存货的处理
			midcun=[]
			midcun.append(u'07')
			midcun.append(bomtitile)
			midcun.append(unicode('单向皮带轮工具','utf-8'))
			midcun.append('')
			midcun.append('')
			midcun.append('')
			ddd=bomtitile.split('-')
			print len(ddd),'dddddddddsdsdsdsddsssssssssssss'
			if len(ddd)==3 and (u'白' in (ddd[2]) or u'黑' in (ddd[2])):
				bomtype='ZNP-'+ddd[1]
				print '111111111111111'
			elif len(ddd)==4:
				bomtype='ZNP-'+ddd[1]+'-'+ddd[2]
				print '22222222222'
			else:
				bomtype=bomtitile
				print '3333333333333'

			midcun.append(bomtitile)
			midcun.append('')
			midcun.append('')
			midcun.append(u'套')
			midcun.append('')
			midcun.append(u'是')
			cun.append(midcun)

			print cun,'cuncunucnu1111111'












		elif u'皮带' in newbom[x][1][1]:
			print newbom[x],'newnenwnebom111111'
			mid=[]
			bomtitile=u'ZNP-'+newbom[x][0][0]
			mid.append(bomtitile)
			mid.append(bomtitile)
			mid.append('')
			mid.append(unicode('单向皮带轮总成BOM','utf-8'))
			mid.append(bomtitile)
			mid.append(unicode('单向皮带轮总成','utf-8'))
			mid.append(u'1')
			
			fux.append(mid)

			#对存货的处理
			midcun=[]
			midcun.append(u'07')
			midcun.append(bomtitile)
			midcun.append(unicode('单向皮带轮总成','utf-8'))
			midcun.append('')
			midcun.append('')
			midcun.append('')
			ddd=bomtitile.split('-')
			print len(ddd),ddd,'dddddddddsdsdsdsddsssssssssssss'
			if len(ddd)==3 and (u'白' in (ddd[2]) or u'黑' in (ddd[2])):
				bomtype=u'ZNP-'+ddd[1]
				print bomtype,'111111111111111'
			elif len(ddd)==4:
				bomtype=u'ZNP-'+ddd[1]+u'-'+ddd[2]
				print bomtype,'22222222222'
			else:
				bomtype=bomtitile
				print bomtype,'3333333333333'

			midcun.append(bomtitile)
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append(u'皮带轮仓库')
			midcun.append(u'装配车间')
			midcun.append('')
			midcun.append(u'套')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')
			midcun.append('')		
			midcun.append(u'是')
			cun.append(midcun)


			print cun,'cuncuncuncucnucn'


	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	sheet1.write(0,0,u'父项模板')
	sheet1.write(1,0,u'BOM编号')
	sheet1.write(1,1,u'BOM名称')
	sheet1.write(1,2,u'类型编码')
	sheet1.write(1,3,u'类型全名')
	sheet1.write(1,4,u'存货编码')
	sheet1.write(1,5,u'存货全名')
	sheet1.write(1,6,u'数量')
	sheet1.write(1,7,u'定额工时')
	sheet1.write(1,8,u'摘要')
	for i in range(len(fux)):
		for j in range (len(fux[i])):
			if is_chinese(fux[i][j]):
				fux[i][j].encode('utf-8')
			# elif not fux[i] and fux[i]!=0: 
			# 	print "空值",
			elif is_num(fux[i][j])==1:
				if math.modf(fux[i][j])[0]==0 or fux[i][j]==0:#获取数字的整数和小数
					fux[i][j]=int(fux[i][j])#将浮点数化成整数
			sheet1.write(i+2,j,fux[i][j])
	book.save('fubom_____2.xls')#存储excel
	book = xlrd.open_workbook('fubom_____2.xls')

				
	book1 = Workbook()
	sheet1 = book1.add_sheet('Sheet 1')
	sheet1.write(0,0,u'基本信息导出模板--存货档案')
	sheet1.write(1,0,u'父项编号')
	sheet1.write(1,1,u'存货编号')
	sheet1.write(1,2,u'存货全名')
	sheet1.write(1,3,u'简名')
	sheet1.write(1,4,u'助记码')
	sheet1.write(1,5,u'规格')
	sheet1.write(1,6,u'型号')
	sheet1.write(1,7,u'产地')
	sheet1.write(1,8,u'条码')
	sheet1.write(1,9,u'缺省供应商')
	sheet1.write(1,10,u'缺省仓库')
	sheet1.write(1,11,u'缺省车间')
	sheet1.write(1,12,u'品牌')
	sheet1.write(1,13,u'基本单位')
	sheet1.write(1,14,u'辅助单位1')
	sheet1.write(1,15,u'单位关系1')
	sheet1.write(1,16,u'辅助单位2')
	sheet1.write(1,17,u'单位关系2')
	sheet1.write(1,18,u'参考零售价')
	sheet1.write(1,19,u'最低售价')
	sheet1.write(1,20,u'一级批发价')
	sheet1.write(1,21,u'二级批发价')
	sheet1.write(1,22,u'三级批发价')
	sheet1.write(1,23,u'四级批发价')
	sheet1.write(1,24,u'五级批发价')
	sheet1.write(1,25,u'六级批发价')
	sheet1.write(1,26,u'七级批发价')
	sheet1.write(1,27,u'八级批发价')
	sheet1.write(1,28,u'九级批发价')
	sheet1.write(1,29,u'十级批发价')
	sheet1.write(1,30,u'十一级批发价')
	sheet1.write(1,31,u'十二级批发价')
	sheet1.write(1,32,u'十三级批发价')
	sheet1.write(1,33,u'十四级批发价')
	sheet1.write(1,34,u'十五级批发价')
	sheet1.write(1,35,u'安全库存天数')
	sheet1.write(1,36,u'备注')
	sheet1.write(1,37,u'副单位')
	sheet1.write(1,38,u'成本计价')
	sheet1.write(1,39,u'安全库存数量')
	sheet1.write(1,40,u'管理批号')
	sheet1.write(1,41,u'近效期先出库')
	sheet1.write(1,42,u'启用序列号管理')
	sheet1.write(1,43,u'是否外购')
	sheet1.write(1,44,u'采购周期')
	sheet1.write(1,45,u'是否自制')
	sheet1.write(1,46,u'生产周期')
	sheet1.write(1,47,u'管理自定义项1')
	sheet1.write(1,48,u'管理自定义项2')
	sheet1.write(1,49,u'管理自定义项3')
	sheet1.write(1,50,u'管理自定义项4')
	sheet1.write(1,51,u'核算方法')
	sheet1.write(1,52,u'参考成本')
	sheet1.write(1,53,u'单位定额成本')
	sheet1.write(1,54,u'存货备用1')
	sheet1.write(1,55,u'存货备用2')
	sheet1.write(1,56,u'存货备用3')
	sheet1.write(1,57,u'存货备用4')
	sheet1.write(1,58,u'存货备用5')
	sheet1.write(1,59,u'存货备用6')
	sheet1.write(1,60,u'存货备用7')
	sheet1.write(1,61,u'存货备用8')
	for i in range(len(cun)):
		for j in range (len(cun[i])):
			if is_chinese(cun[i][j]):
				cun[i][j].encode('utf-8')
			# elif not cun[i] and cun[i]!=0: 
			# 	print "空值",
			elif is_num(cun[i][j])==1:
				if math.modf(cun[i][j])[0]==0 or cun[i][j]==0:#获取数字的整数和小数
					cun[i][j]=int(cun[i][j])#将浮点数化成整数
			sheet1.write(i+2,j,cun[i][j])
	book1.save('cunhuo_____1.xls')#存储excel
	book1 = xlrd.open_workbook('cunhuo_____1.xls')




















def out():
	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	global line
	line=0
	for i in range(len(newbom)):
		# if newbom[i]
	
		for j in range (len(newbom[i])):

			if len(newbom[i][j])==1:
				if is_chinese(newbom[i][j]):
					newbom[i][j].encode('utf-8')
				# elif not newbom[i] and newbom[i]!=0: 
				# 	print "空值",
				elif is_num(newbom[i][j])==1:
					if math.modf(newbom[i][j])[0]==0 or newbom[i][j]==0:#获取数字的整数和小数
						newbom[i][j]=int(newbom[i][j])#将浮点数化成整数
				# print i,j,newbom[i][j],'?????????????'
				sheet1.write(line,j,newbom[i][j])
				line=line+1
			else:
				for z in range(len(newbom[i][j])):
					if is_chinese(newbom[i][j][z]):
						newbom[i][j][z].encode('utf-8')
					# elif not newbom[i] and newbom[i]!=0: 
					# 	print "空值",
					elif is_num(newbom[i][j][z])==1:
						if math.modf(newbom[i][j][z])[0]==0 or newbom[i][j][z]==0:#获取数字的整数和小数
							newbom[i][j][z]=int(newbom[i][j][z])#将浮点数化成整数
					# print line,z,'!!!!!!!!!!!!'
					sheet1.write(line,z,newbom[i][j][z])
				line=line+1
	book.save('finalinput_bom.xls')#存储excel
	book = xlrd.open_workbook('finalinput_bom.xls')
					




if __name__ == "__main__":
	global bom,newbom,fux,zix
	bom=[]
	newbom=[]
	fux=[]
	zix=[]
	readexcel('BOM')
	# fu()
	# zi()
	# checkall('all')
  