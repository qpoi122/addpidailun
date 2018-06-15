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
		changeflag=0
		b=[]
		lack=[u'皮带轮毂',u'外曲面芯轮',u'六角螺纹压件',u'轴承连接器',u'螺纹芯轴']
			
		table = workbook.sheets()[name]
		# print "第",name+1,"个sheet："
		#获取所有的行数
		nrows=table.nrows
		# print nrows,'nrows'
		# if not nrows:
		# 	break
		title=[]
		title.append(table.name)
		# print title,'11111'
		if title[0]!='Sheet1':
			b.append(title)
			# print title,'2222222222222222222222'

			for n in range(nrows):
				a=table.row_values(n)
				c=[]
				if a[0]:
					for l in range(len(a)):
						if l==6 and a[l]:
							c.append(unicode('自制件','utf-8'))
						elif l==7 and a[l]:
							c.append(unicode('外购件','utf-8'))
						else :c.append(a[l])
					b.append(c)

		#处理没有表明类型的
		for x in range(len(b)):
			if len (b[x])>1 and b[x][6]=='' and b[x][7]=='' and b[x][1] in lack:
				b[x][6]=u'自制件'
				print '6666666666666666666666',b[0]
			elif len (b[x])>1 and b[x][6]=='' and b[x][7]=='' and b[x][1] not in lack:
				b[x][7]=u'外购件'
				print '777777777777777777777777',b[0]
		if len(b)>1:
			allitem.append(b)
		

	print allitem,'alllllllllllllllllllll'


def readnew(content):
	filename(content)
	#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		# print "第",name+1,"个sheet："
		#获取所有的行数
		nrows=table.nrows
		for n in range(nrows):
			a=table.row_values(n)
			b=table.row_values(0)
			c=[]
			d=[]#新增的数组
			for l in range(len(a)):
				c.append(a[l])
			c.append(table.name)
			newitem.append(c)
	print newitem,'newewewewe'
	changeline=[]
	for x in range(len(allitem)):
		for y in range(len(allitem[x])):
			if len(allitem[x][y])>1:
				for z in range(len(newitem)):
					if allitem[x][y][1] in [u'皮带轮毂',u'外曲面芯轮'] and len(allitem[x][y][0].split("-"))==3:
						need='ZP-'+allitem[x][y][0].split("-")[1]
					else:
						need=allitem[x][y][0]

					if need==newitem[z][0] and u'保护盖' not in allitem[x][y][1]:
						allitem[x][y][0]=newitem[z][1]
						changeline.append(allitem[x][0])
						changeline.append(allitem[x][y][0])
	# print allitem,'last'
	print changeline,'changeline'


def specaldeel(content):
	filename(content)
	spchange=[]
	Sheetname=workbook.sheet_names()
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]		
		#获取所有的行数
		nrows=table.nrows
		mid=[]
		for n in range(nrows):
			c=[]

			a=table.row_values(n)
			# c.append(table.name)
			for l in range(len(a)):
				c.append(a[l])
			mid.append(c)
		specal.append(mid)
	print specal,'specaldeel'

	print len(specal[0]),len(specal[1])
	for z in range (len(specal[1])):
		print specal[1][z],'15315313515315315'

	print allitem,'alitemaltieam'

	for x in range(len(allitem)):
		for j in range(len(specal)):
			print allitem[x][2][0],'alllllllllllllllllllllllllllllll'
			if 'ZP-46' in allitem[x][2][0] or 'ZP-47' in allitem[x][2][0]:
				for y in range(1,len(allitem[x])):
					for z in range (len(specal[0])):
						if allitem[x][y][0]==specal[0][z][0]:
							allitem[x][y][0]=specal[0][z][1]
							allitem[x][y][1]=specal[0][z][2]
							spchange.append(allitem[x][0][0])
							spchange.append(allitem[x][y][0])
							
			elif 'ZP-48' in allitem[x][2][0]:
				flag=0
				for y in range(1,len(allitem[x])):
					if y >=len(allitem[x]):
						break
					for z in range (len(specal[1])):
						if allitem[x][y][1]==u'卡簧':
							del allitem[x][y]
							flag=1
						if flag!=1:
							if allitem[x][y][0]==specal[1][z][0]:
								allitem[x][y][0]=specal[1][z][1]
								allitem[x][y][1]=specal[1][z][2]
								spchange.append(allitem[x][0][0])
								spchange.append(allitem[x][y][0])	
						else:
							if allitem[x][y-1][0]==specal[1][z][0]:
								allitem[x][y-1][0]=specal[1][z][1]
								allitem[x][y-1][1]=specal[1][z][2]
								spchange.append(allitem[x][0][0])
								spchange.append(allitem[x][y-1][0])
								flag=0

			elif 'ZP-49' in allitem[x][2][0]:
				for y in range(1,len(allitem[x])):
					for z in range (len(specal[2])):
						if allitem[x][y][0]==specal[2][z][0]:
							allitem[x][y][0]=specal[2][z][1]
							allitem[x][y][1]=specal[2][z][2]
							spchange.append(allitem[x][0][0])
							spchange.append(allitem[x][y][0])

			if 'E' in allitem[x][0][0]:
				for y in range(1,len(allitem[x])):
					print allitem[x][2][0],allitem[x][y][0],'??????????????'
					if allitem[x][y][0]==6905.0:
						allitem[x][y][0]=u'6905-NTN'
						print allitem[x][y],'tntntntntnttntntntn'

	print spchange,'spchage'



def out():
	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	global line
	line=0
	for i in range(len(allitem)):
		# if allitem[i]
	
		for j in range (len(allitem[i])):

			if len(allitem[i][j])==1:
				if is_chinese(allitem[i][j]):
					allitem[i][j].encode('utf-8')
				# elif not allitem[i] and allitem[i]!=0: 
				# 	print "空值",
				elif is_num(allitem[i][j])==1:
					if math.modf(allitem[i][j])[0]==0 or allitem[i][j]==0:#获取数字的整数和小数
						allitem[i][j]=int(allitem[i][j])#将浮点数化成整数
				# print i,j,allitem[i][j],'?????????????'
				sheet1.write(line,j,allitem[i][j])
				line=line+1
			else:
				for z in range(len(allitem[i][j])):
					if is_chinese(allitem[i][j][z]):
						allitem[i][j][z].encode('utf-8')
					# elif not allitem[i] and allitem[i]!=0: 
					# 	print "空值",
					elif is_num(allitem[i][j][z])==1:
						if math.modf(allitem[i][j][z])[0]==0 or allitem[i][j][z]==0:#获取数字的整数和小数
							allitem[i][j][z]=int(allitem[i][j][z])#将浮点数化成整数
					# print line,z,'!!!!!!!!!!!!'
					sheet1.write(line,z,allitem[i][j][z])
				line=line+1
	book.save('PiDaiLun_BOM.xls')#存储excel
	book = xlrd.open_workbook('PiDaiLun_BOM.xls')



def checkall(content):
	# filename(content)
	# Sheetname=workbook.sheet_names()
	# for name in range(len(Sheetname)):
	# 	table = workbook.sheets()[name]
	# 	nrows=table.nrows
	# 	for n in range(nrows):
	# 		mid=[]
	# 		a=table.row_values(n)
	# 		for i in range(len(a)):
	# 			if is_chinese(a[i]):
	# 				a[i]=a[i].encode('utf-8' )
	# 			elif is_num(a[i])==1:
	# 				if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
	# 					a[i]=int(a[i])#将浮点数化成整数
	# 			mid.append(a[i])
	# 		lingjian.append(mid)
	# print lingjian,'lingjian'


	filename(content)
	Sheetname=workbook.sheet_names()
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]
			a=table.row_values(n)
			if len(a)>2:
				for i in range(1,3):
					if is_chinese(a[i]):
						a[i]=a[i].encode('utf-8' )
					elif is_num(a[i])==1:
						if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
							a[i]=int(a[i])#将浮点数化成整数
					mid.append(a[i])
			inputed.append(mid)
	print inputed,'inputed'


	for x in range(len(allitem)):
		for y in range(len(allitem[x])):
			flag=0
			if len(allitem[x][y])>1:
				for z in range (len(inputed)):			
					if allitem[x][y][0] == inputed[z][0]:
						allitem[x][y][1]=inputed[z][1]
						flag=1
						break
				if allitem[x][y][1] in [u'单向皮带轮']:
					flag=1

				if flag==0:
					cun=0
					for d in range(len(needinput)):
						if allitem[x][y][0]==needinput[d][0]:
							cun=1
							break
					if cun==0:
						c=[]
						c.append(allitem[x][y][0])
						c.append(allitem[x][y][1])
						needinput.append(c)




# #没有匹配到的
# 	for x in range(len(lingjian)):
# 		cunzai=0
# 		for y in range(len(inputed)):
# 			if type(lingjian[x][2]) is types.IntType :
# 				# print type(lingjian[x][2]),lingjian[x][2],'sdsadasdads'
# 				if lingjian[x][0] == inputed[y][0]:
# 					cunzai=1
# 					break
# 			else:
# 				cunzai=1
# 		if cunzai==0:
# 			if lingjian[x] not in needinput:
# 				needinput.append(lingjian[x])



# 	#处理型号一样但是名字不一样的情况
	
# 	for x in range (len(lingjian)):
# 		for y in range(len(inputed)):
# 			if type(lingjian[x][2]) is types.IntType:
# 				if lingjian[x][1]!=inputed[y][1] and lingjian[x][0]==inputed[y][0]:
# 					lingjian[x][1]=inputed[y][1]


	# book = Workbook()
	# sheet1 = book.add_sheet('Sheet 1')
	# for n in range(len(needinput)):#将相符的内容显示出来
	# 	for i in range(len(needinput[n])):#数据逐行写入excel
	# 		# print len(gett[n])
	# 		if is_chinese(needinput[n][i]):
	# 			needinput[n][i].encode('utf-8')
	# 		elif is_num(needinput[n][i])==1:
	# 			if math.modf(needinput[n][i])[0]==0 or needinput[n][i]==0:#获取数字的整数和小数
	# 				needinput[n][i]=int(needinput[n][i])#将浮点数化成整数
	# 		sheet1.write(n,i,needinput[n][i])
	# book.save('needinput.xls')#存储excel
	# book = xlrd.open_workbook('needinput.xls')



	book1= Workbook()
	sheet1 = book1.add_sheet('Sheet 1')
	global line
	line=0
	for i in range(len(allitem)):
		# if allitem[i]
	
		for j in range (len(allitem[i])):

			if len(allitem[i][j])==1:
				if is_chinese(allitem[i][j]):
					allitem[i][j].encode('utf-8')
				# elif not allitem[i] and allitem[i]!=0: 
				# 	print "空值",
				elif is_num(allitem[i][j])==1:
					if math.modf(allitem[i][j])[0]==0 or allitem[i][j]==0:#获取数字的整数和小数
						allitem[i][j]=int(allitem[i][j])#将浮点数化成整数
				# print i,j,allitem[i][j],'?????????????'
				sheet1.write(line,j,allitem[i][j])
				line=line+1
			else:
				for z in range(len(allitem[i][j])):
					if is_chinese(allitem[i][j][z]):
						allitem[i][j][z].encode('utf-8')
					# elif not allitem[i] and allitem[i]!=0: 
					# 	print "空值",
					elif is_num(allitem[i][j][z])==1:
						if math.modf(allitem[i][j][z])[0]==0 or allitem[i][j][z]==0:#获取数字的整数和小数
							allitem[i][j][z]=int(allitem[i][j][z])#将浮点数化成整数
					# print line,z,'!!!!!!!!!!!!'
					sheet1.write(line,z,allitem[i][j][z])
				line=line+1
	book1.save('BOM.xls')#存储excel
	book1 = xlrd.open_workbook('BOM.xls')




def peijian(content):
	filename(content)
	pop=[]
	notright=[]
	Sheetname=workbook.sheet_names()
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(nrows):
			mid=[]
			a=table.row_values(n)
			for i in range(len(a)):
				if is_chinese(a[i]):
					a[i]=a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				mid.append(a[i])
			if mid[1]!='':
				# if mid[0] not in pop and u'皮带轮毂' not in mid[1] and u'工具' not in mid[1] and u'外曲面芯轮' not in mid[1] :
				if mid[0] not in pop :
					pop.append(mid[0])
					pop.append(mid[1])
					print  mid[0]
					if mid[1] in [u'皮带轮毂',u'外曲面芯轮'] :
						if len(mid[0].split("-")[1])<5 or (len(mid[0].split("-")[1])==5 and 'A' in mid[0]) or (len(mid[0].split("-")[1])==5 and u'特' in mid[0]):
							notright.append(mid[0])
							notright.append(mid[1])
			suoyoupeijian.append(mid)
	# print suoyoupeijian,'suoyoupeijian'
	print pop,'pop'

	book1 = Workbook()
	sheet1 = book1.add_sheet('Sheet 1')
	line=0
	for n in range(len(pop)):#将相符的内容显示出来
		if is_chinese(pop[n]):
			pop[n].encode('utf-8')
		elif is_num(pop[n])==1:
			if math.modf(pop[n])[0]==0 or pop[n]==0:#获取数字的整数和小数
				pop[n]=int(pop[n])#将浮点数化成整数
		sheet1.write(int(n//2),line,pop[n])
		if line==0:
			line=1
		elif line==1:
			line=0

	book1.save('pop.xls')#存储	excel
	book1 = xlrd.open_workbook('pop.xls')

	# book2 = Workbook()
	# sheet1 = book2.add_sheet('Sheet 1')
	# line=0
	# for n in range(len(notright)):#将相符的内容显示出来
	# 	if is_chinese(notright[n]):
	# 		notright[n].encode('utf-8')
	# 	elif is_num(notright[n])==1:
	# 		if math.modf(notright[n])[0]==0 or notright[n]==0:#获取数字的整数和小数
	# 			notright[n]=int(notright[n])#将浮点数化成整数
	# 	sheet1.write(int(n//2),line,notright[n])
	# 	if line==0:
	# 		line=1
	# 	elif line==1:
	# 		line=0

	# book2.save('notright.xls')#存储excel
	# book2 = xlrd.open_workbook('notright.xls')




if __name__ == "__main__":
	global allitem,newitem,specal,lingjian,inputed,needinput,suoyoupeijian
	allitem=[]
	newitem=[]
	specal=[]
	lingjian=[]
	inputed=[]
	needinput=[]
	suoyoupeijian=[]
	readexcel('addnew')
	# readnew('new')
	# specaldeel('change')
	out()
	peijian('PiDaiLun_BOM')
	checkall('存货档案')
	# readnew2('new2')

