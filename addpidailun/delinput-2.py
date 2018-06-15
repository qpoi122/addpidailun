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
import time  
import pandas

from datetime import datetime
from xlrd import xldate_as_tuple
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








#读取帐目，生成的是标题和内容一对一的
def readexcel(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		nrows=table.nrows
		title=table.row_values(1)
		print title,'titititiitittiitititiitit'
#将起始行所有的职工名字获取起来,起始行应包括所有人员
		# pd=[4,6,8,10,12]
		# pd=[u'客户编号',u'总成编号',u'数量',u'日期',u'塑料袋',u'小标贴',u'小内盒']
		pd=[u'日期',u'塑料袋']
		notmsag=[]
		ren=[]
		for x in range (len(title)):#根据信息将内容切片		
			for y in pd:
				if y==title[x]:
					notmsag.append(x)
					break
		print notmsag,'notmsag'
		for y in range(len(notmsag)-1):
			for x in range (len(title)):
				ren.append(title[notmsag[y]+1:notmsag[y+1]])
				break
		print ren,'renrnenrnernenrnenrnenrn'
			# print pd[x]
			# print title[pd[x]],'titiitit'
			# print len(title)
			# print title
			# print title[pd[x]],'dddd'



	
		for x in range(len(ren)):
			mid=[]
			for y in range(len(ren[x])):
				if u'/' in ren[x][y]:
					aa=ren[x][y].split(u'/')
					for z in aa:
						if z not in mid:
							mid.append(z)

				elif u'和' in ren[x][y]:
					aa=ren[x][y].split(u'和')
					for z in aa:
						if z not in mid:
							mid.append(z)
				else:
					if ren[x][y] not in mid:
						mid.append(ren[x][y])
			renming.append(mid)

		print renming,'renrnenrnenrnenrnenrnennrnenrnenrnn'
					
		# for x in range (len(title)):			
		# 	ren=[]
		# 	flag=0
	

		# 	if u'/' in title[x]:
		# 		ren=title[x].split(u'/')
		# 	elif u'和' in title[x]:
		# 		ren=title[x].split(u'和')
		# 	else:
		# 		for z in pd:
		# 			if z ==title[x]:
		# 				flag=1
		# 				break
		# 		if flag==0:
		# 			ren=title[x]
		# 		else:
		# 			ren=[]
					
		
		# 	mid=[]
		# 	# mid.append(x)
		# 	if ren!=[]:
		# 		for y in ren:
		# 			mid.append(y)
		# 		renming.append(mid)
		# print renming,'renrenrnenrnernenr'
#第1和第二列都为第一道做的修改
		renming[0]=renming[0]+renming[1]
		del renming[1]
		print renming,'renrenrnenrnernenr'

#获取每行的信息加上抬头
		# print title,'title'
		for n in range(2,nrows):			
		#获取每行内容
			a=table.row_values(n)
			mid=[]
			for i in range(len(a)):	


				if is_chinese(a[i]):
					a[i].encode('utf-8')
					
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				try:
					a[i]=a[i].strip()
				except:
					pass

	

			
				mid.append(title[i])
				mid.append(a[i])
			# print mid,'middmimdi'
			#如果某行没有总数量，则找到非空行，更改抬头
			# if mid[3]=='':
			# 	panduan=[9,13,17,21,25]
			# 	for x in range(len(panduan)):
			# 		if mid[(panduan[x])]!='':
			# 			title[int((panduan[x]-1)/2)]=mid[(panduan[x])]

			# 			ren=[]
						
			# 			# ren.append(x)
			# 			if u'/' in mid[(panduan[x])]:
			# 				ren=mid[(panduan[x])].split(u'/')
			# 			elif u'和' in mid[(panduan[x])]:
			# 				ren=mid[(panduan[x])].split(u'和')
			# 			else:
			# 				ren.append(mid[(panduan[x])])
			# 			# print ren ,'2222222222222222'
			# 			flag=0

			# 			#除了第一排匹配到补充人数
			# 			# for  y in ren:
			# 			# 	print ren ,renming[x],'!!!!!!!!!!!!!!!'


			# 			# 	# print y ,ren,'ddddddddddd'
			# 			# 	if y not in renming[x]:
			# 			# 		renming[x].append(y)
				



			
			
			# else:
				# if mid not in zhangdan:
			zhangdan.append(mid)
			# print renming,'renrenrnenrenrss'
	# print renming,'sdsdsdsd'
	# print zhangdan,'zhangdan'
	# diyi=0
	# for x in range(len(zhangdan)):
	# 	diyi=diyi+zhangdan[x][u'第一']
	# print diyi,'yi'

	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	for i in range(len(zhangdan)):
		for j in range (len(zhangdan[i])):
			if is_chinese(zhangdan[i][j]):
				zhangdan[i][j].encode('utf-8')
			# elif not zhangdan[i] and zhangdan[i]!=0: 
			# 	print "空值",
			elif is_num(zhangdan[i][j])==1:
				if math.modf(zhangdan[i][j])[0]==0 or zhangdan[i][j]==0:#获取数字的整数和小数
					zhangdan[i][j]=int(zhangdan[i][j])#将浮点数化成整数
			sheet1.write(i,j,zhangdan[i][j])
	book.save('allmesg.xls')#存储excel
	book = xlrd.open_workbook('allmesg.xls')

def readexcel1(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		
		table = workbook.sheets()[name]
		ttype=table.name
		nrows=table.nrows
		for n in range(nrows):			
		#获取每行内容
			a=table.row_values(n)
			mid=[]
			mid.append(ttype)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数 
				mid.append(a[i])
			money.append(mid)

	# print money,'montytyyt'


# def addsi(zhangdan,money):


def readexcel2(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		
		table = workbook.sheets()[name]
		ttype=table.name
		nrows=table.nrows
		midd=[]
		midd.append(ttype)
		for n in range(nrows):			
		#获取每行内容
			a=table.row_values(n)
			mid=[]
	
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数 
				mid.append(a[i])
			midd.append(mid)
		taizhang.append(midd)
	print taizhang,'taizhang'


# def addsi(zhangdan,money):

# def new_
def addname():
    # n=1
    # while True:
    #     yield n
    #     n+=2yield n
    for x in renming:
		aa=dict.fromkeys(x,[])
		yield aa

# def new_
def addnewname():
    # n=1
    # while True:
    #     yield n
    #     n+=2yield n
    for x in smulu:
		aa=dict.fromkeys(x,[])
		yield aa





	
if __name__ == "__main__":
	global zhangdan,money,gongzi,errortype,renming,needi,taizhang
	zhangdan=[]
	money=[]
	gongzi=[]
	errortype=[]
	renming=[]
	needi=[]
	taizhang=[]
	readexcel('12345')
	readexcel1('newmy')
	readexcel2('taizhang')


	add_name=addname()
	for o in add_name:
		needi.append(o)
	print needi,'needi' 

#根据获取到的人名,再次遍历账单,将碰到的信息分配给对应的人下面,并且去比对money表加上钱

	for x in zhangdan:
		for z in range(len(x)):
			#提前获取各个信息
			pd=[u'客户编号',u'总成型号',u'数量',u'日期',u'塑料袋',u'小标贴',u'小内盒']
			
			# zcbh=[]
			# khbh=[]
			# print x,'xxxxxxxxxxxxxxxxxxxxx'
			if x[z]==u'总成型号':
				# print 'qioeuoqidoaih'
				zcbh=x[z+1]
			elif  x[z]==u'客户编号':
				khbh=x[z+1]
			elif  x[z]==u'塑料袋':
				sld=x[z+1]
			elif  x[z]==u'小标贴':
				xbt=x[z+1]
			elif  x[z]==u'小内盒':
				xnh=x[z+1] 
	
		for z in range(len(x)):
			pipei_num=[]
			if x[z] not in pd and z%2==0:
				try:
					if u'/' in x[z]:
						pipei_num=x[z].split(u'/')
					elif u'和' in x[z]:
						pipei_num=x[z].split(u'和')
					else:
						pipei_num.append(x[z])
				except:
					pass

			for m in range(len(pipei_num)):
			#将a中匹配到的键，值取出，然后增加值，然后放回去
				for h in range(len(needi)):	
					if needi[h].has_key(pipei_num[m]) and x[z+1]!=u'' and x[z+1]!=u' ' and type(x[z+1])!=type(u''):
						zhongji=[]#zong 
						mid=[]#xiao 
						zhongji=copy.deepcopy(needi[h].get(pipei_num[m]))						
				
						flgggggg=0
						a=z-1
						for p in range(0,3):
							# print x[a],'wdwdwdwdwdwd'
							if type(x[a])==type(1):
								mid.append(x[a])
								flgggggg=1
								break
							else:
								a=a-1
								# print '11111'
						if flgggggg==0:
							mid.append(u'')
						# mid.append(x[1])

						mid.append(zcbh)
						mid.append(khbh)
						if type(x[z+1])!=type(1): 
							print z,x[z+1],type(x[z+1]),type(len(pipei_num))
							print z,x[z+1],type(x[z+1]),type(len(pipei_num))
							print z,x[z+1],type(x[z+1]),type(len(pipei_num))
						#这里可能应为汉字报错


# #划分数量
# 						shul=x[z+1]/len(pipei_num)
# 						mid.append(shul)

# 						zongmoney=0
# 						finded=0
# 						for j in range(len(money)):
# 							# if h!=0:
# 							# 	print int(money[j][0]),h,type(int(money[j][0])),type(h)
# 								if int(money[j][0])==h+1 and money[j][1]==x[3] and int(money[j][0])!=4:
# 									# print 'ssssssssssssssssss',money[j][2],shul,type(money[j][2])
# 									try:
# 										zongmoney=money[j][2]*shul
# 									except:
# 										pass
# 									# print zongmoney
# 									break

# 						# if h+1==4 and finded==0:
# 						# 	for j in range(len(money)):
# 						# 		flag=0
# 						# 		if zhan
						mid.append(len(pipei_num))
#划分单价	
						shul=x[z+1]
						mid.append(shul)

						zongmoney=0
						finded=0
						for j in range(len(money)):
							# if h!=0:
							# 	print int(money[j][0]),h,type(int(money[j][0])),type(h)

							#将除了第四道的工序都匹配金钱,没有匹配到的钱为0
								if int(money[j][0])==h+1 and money[j][1]==x[3] and int(money[j][0])!=4:
									# print 'ssssssssssssssssss',money[j][2],shul,type(money[j][2])
									try:
										zongmoney=money[j][2]*shul/len(pipei_num)
									except:
										pass
									# print zongmoney
									break



						mid.append(zongmoney)
						#如果是第四种加上√ 信息
						if h==3: 
							# mid.append(x[23])
							# mid.append(x[25])
							# mid.append(x[27])

							mid.append(sld)
							mid.append(xbt)
							mid.append(xnh)
						zhongji.append(mid)
						needi[h][pipei_num[m]]=zhongji
						# print zhongji,needi[h]




	print needi,'djdjdjdj'

	smulu=[]
	for x in needi:
		dlist=list(x.keys())
		smulu=smulu+dlist
	print smulu	,'shushuhushushushushu'


	print taizhang,'sdsdsdswwwwwwwwwww'

#核对一下账目和台账是否匹配
	lackzm=[]
	add_name1=addnewname()

	for x in 
	# for o in add_name1:
	lackzm.append(add_name1)
	print lackzm,'lackcsdsdsdsdsd' 

	
	lacktj=[]
	for x in smulu:
		zmsj=[]
		for y in needi:

			if y.has_key(x):
				zmsj=copy.deepcopy(y.get(x))	
				break
		tj=[]
		for z in range(len(taizhang)):
			print taizhang[z][0],'what the fuck'
			if x in taizhang[z][0]:
				tj=copy.deepcopy(taizhang[z])
				break
		for m in range(len(zmsj)):
			flag=0
			for n in range(1,len(tj)):
				if zmsj[m][1]==tj[n][1] and zmsj[m][2]==tj[n][2] and zmsj[m][4]==tj[n][3] :
					# if zmsj[m][0]-tj[n][0]<=5 and zmsj[m][0]-tj[n][0]>=-5:
					flag=1
					zmsj[m].append(u'o')
					tj[n].append(u'o')
					break

			if flag==0:
				zmsj[m].append(u'x')

		for mm in lackzm:
			if mm.has_key(x): 
				dd=mm.get(x)
				for k in zmsj:
					dd.append(k)


	print lackzm,'sdhushdiushidhsihdisuhdiushdiu'


	book1 = Workbook()
	for x in smulu:
		sheet1 = book1.add_sheet(x)
		for y in lackzm:
			if y.has_key(x):
				sss=copy.deepcopy(y.get(x))
				for i in range(len(sss)):
					for j in range (len(sss[i])):
						# print 'xxxxx'
						if is_chinese(sss[i][j]):

								sss[i][j].encode('utf-8')

						# elif not sss[i] and sss[i]!=0: 
						# 	print "空值",
						elif is_num(sss[i][j])==1:
							if math.modf(sss[i][j])[0]==0 or sss[i][j]==0:#获取数字的整数和小数
								sss[i][j]=int(sss[i][j])#将浮点数化成整数
						
						# print sss[i][j]
						if j==0:
							try:
								cccc=pandas.to_datetime(sss[i][j]-25569,unit='d')
								sss[i][j]=str(pandas.Period(cccc, freq='D'))
							except:
								pass
						sheet1.write(i,j,sss[i][j])
	book1.save('lackzm.xls')#存储excel
	book1 = xlrd.open_workbook('lackzm.xls')










#对第四道进行单独处理,先匹配名字一样的,没有再去按老方法处理


	for x in smulu:
		if needi[3].has_key(x):
			# print '111'
			for y in needi[3].get(x):
				


				# print '2222'
				for z in range(len(money)):
					# print '3333'
					if money[z][0]==u'4' and y[1]==money[z][1]:
						# print y,y[6],'y6'
						duigeshu=len(y[7])
						jiage=((duigeshu-1)*0.005+money[z][2])/y[3]
						print y[4],'2222222'
						y[5]=jiage*y[4]
						# print jiage,y[3],y[1],y[4],'1111111111111'



	for x in smulu:
		if needi[3].has_key(x):
			# print '111'
			for y in needi[3].get(x):
				flag=0
				for z in range(len(money)):
					# print '3333'
					if money[z][0]==u'5' and  y[5]==0:
						# print '4444444444'
						#内盒一样并且带-的
						# print len(y),len(money[z]),(y),money[z]
						# print y
						if y[8]==money[z][3] and u'-'==money[z][1] and type(y[1])==type(u'a') :
							# print '11111111111111111111111111111'
							# if (type(zhangdan[x][3])==type(u'a') and len(zhangdan[x][3])>6):
							jiage=money[z][4]/y[3]
							y[5]=jiage*y[4]
							flag=1
							# print y[4],'y3yy33yy3y3y3y3'
							break
						#所有带几盒几个的,但是不带-的
						elif u'个'  in y[7] or u'只'  in y[7]:
							if (type(y[1])==type(u'a') and u'-' not in y[1]) or type(y[1])==type(1):
							
								if y[6]==money[z][1]  and y[8]==money[z][3]:
									# print '1111111111111111111111111111'
									if y[7]=='' and y[7]==money[z][2]:
										# print '222222222222222222222222222222222222222'
										jiage=money[z][4]/y[3]
										y[5]=jiage*y[4]
										flag=1
										break
									elif y[7]!='' and money[z][2]==u'√':
										geshu=len(y[7])
										jiage=((geshu-1)*0.005+money[z][4])/y[3]
										y[5]=jiage*y[4]
										flag=1
										break
						#带排/数的			
						elif  (type(y[1])==type(u'a') and u'-' not in y[1] and (y[8]==u'排' or y[8]==u'数')) or (type(y[1])==type(1) and (y[8]==u'排' or y[8]==u'数')):
							# print type(y[1])==type(u'a'),y[1],'sdsadsadsdasd'
							s=[u'A',u'B',u'C',u'D',u'E',u'F',u'G',u'H']
							youzimu=0	
							lenxinhao=0
							if type(y[1])==type(u'a'):
								for p in range(len(s)):
									if s[p] in y[1]:
										youzimu=1
										break

							if youzimu==1:
								lenxinhao=len(y[1])-1
							else:
								try:
									lenxinhao=len(y[1])
								except:
									if y[1]>10000 and y[1]<30000:
										lenxinhao=5
									else:
										lenxinhao=4

							if lenxinhao==money[z][1] and y[8]==money[z][3]:
								jiage=money[z][4]/y[3]
								y[5]=jiage*y[4]
								flag=1
								break

						#剩余的正常型号
						elif  y[6]==money[z][1]  and y[8]==money[z][3]:
							# print y[1],'SUSUSUSUSSU'
							if  y[7]=='' and y[7]==money[z][2]:
								jiage=money[z][4]/y[3]
								y[5]=jiage*y[4]
								flag=1
								break
							elif y[7]!='' and money[z][2]==u'√':
								geshu=len(y[7])
								jiage=((geshu-1)*0.005+money[z][4])/y[3]
								y[5]=jiage*y[4]
								flag=1
								# print y[1],jiage,geshu,money[z][4],'2323232323'
								break
					
					# else:
					# 	flag1=1			
					# 	mid=[]
					# 	mid.append(zhangdan[x][5])
					# 	mid.append(zhangdan[x][3])
					# 	mid.append(u'4')
					# 	errortype.append(mid)
					# 	break	


	# print needi,'sssssssssssss'



	#补差价
	tijia=[]
	for x in range(len(money)):
		if money[x][0]==u'6':
			
			tijia.append(money[x])
	print tijia,'titititi'
	for x in range(len(tijia)):
		print renming[tijia[x][1]-1],'11111111111111111bbbbbbbbbbb'
		for y in renming[tijia[x][1]-1]:
			for z in needi[tijia[x][1]-1].get(y):
				# print z[4],tijia[x][1],'??????????????'
				if z[4]<=tijia[x][2]:
					# print 'ddddddddddddddddddd'
					newjiage=z[5]*(1+tijia[x][3])
					# print newjiage
					z.append(newjiage)
					# print z
				else:
					z.append(z[5])

	book = Workbook()
	for x in smulu:
		sheet1 = book.add_sheet(x)
		for y in needi:
			if y.has_key(x):
				sss=copy.deepcopy(y.get(x))
				for i in range(len(sss)):
					for j in range (len(sss[i])):
						# print 'xxxxx'
						if is_chinese(sss[i][j]):

								sss[i][j].encode('utf-8')

						# elif not sss[i] and sss[i]!=0: 
						# 	print "空值",
						elif is_num(sss[i][j])==1:
							if math.modf(sss[i][j])[0]==0 or sss[i][j]==0:#获取数字的整数和小数
								sss[i][j]=int(sss[i][j])#将浮点数化成整数
						
						# print sss[i][j]
						if j==0:
							try:
								cccc=pandas.to_datetime(sss[i][j]-25569,unit='d')
								sss[i][j]=str(pandas.Period(cccc, freq='D'))
							except:
								pass
						sheet1.write(i,j,sss[i][j])
	book.save('suoyou.xls')#存储excel
	book = xlrd.open_workbook('suoyou.xls')

	# addmoney()

	# print yi,er,san,si
	# print yilist,'yi',len(yilist)
	# print erlist,'er',len(erlist)
	# print sanlist,'san',len(sanlist)
	# print silist,'si',len(silist)
	# print errortype,'errortype'