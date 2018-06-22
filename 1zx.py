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
######
reload(sys)
sys.setdefaultencoding( "utf-8" )

#最终处理
def lastgett(gett):
	fl=0
	nonetype=[]
	threepin=[]
	twopin=[]
	fivepin=[]
	fourpin=[]
	zhuangwan=[]
	pi=[]
	pintype=[]
	chilun=[]
	lasttype=[]
	duzi=[]
	for x in range (len(gett)):
		if gett[x][gett[x].index('每箱规格')+1]=='none':
			nonetype.append(gett[x])
		elif gett[x][gett[x].index('每箱规格')+1].find('拼')!=-1:
				pintype.append(gett[x])
		elif gett[x][gett[x].index('每箱规格')+1].find('最终')!=-1 :
			duzi.append(gett[x])
		elif gett[x][gett[x].index('每箱规格')+1].find('全')!=-1 :
			zhuangwan.append(gett[x])
		elif gett[x][gett[x].index('品名')+1].find('皮带轮')!=-1:
				pi.append(gett[x])
		elif gett[x][gett[x].index('品名')+1].find('齿轮头')!=-1:
				chilun.append(gett[x])
		else: lasttype.append(gett[x])




	print nonetype,'nonetype'
	print lasttype,'lasttype'
	print pintype,'pintype'

	for x in range(0,len(pintype)):
		print pintype[x][pintype[x].index('每箱规格')+1][0],'xxxxwqwrqfsdgbsfdg'
		if int(pintype[x][pintype[x].index('每箱规格')+1][0])==3:
			threepin.append(pintype[x])
			print '???threepin',threepin
		elif int(pintype[x][pintype[x].index('每箱规格')+1][0])==2:
			twopin.append(pintype[x])
			print '???twopin',twopin
		elif int(pintype[x][pintype[x].index('每箱规格')+1][0])==5:
			fivepin.append(pintype[x])
			print '???twopin',fivepin
		elif int(pintype[x][pintype[x].index('每箱规格')+1][0])==4:
			fourpin.append(pintype[x])
			print '???twopin',fourpin

	print threepin,'threepin'
	print twopin,'twopin'
	print fivepin,'fivepin'
	print fourpin,'fourpin'

	maxfive=0
	for x in range(0,len(fivepin)):
		print int(fivepin[x][fivepin[x].index('每箱规格')+1].split("CM")[1]),'ththththtetewrtwerwer'
		if int(fivepin[x][fivepin[x].index('每箱规格')+1].split("CM")[1])>maxfive:
			maxfive=int(fivepin[x][fivepin[x].index('每箱规格')+1].split("CM")[1])

	print maxfive,'maxfive'

	maxfour=0
	for x in range(0,len(fourpin)):
		print int(fourpin[x][fourpin[x].index('每箱规格')+1].split("CM")[1]),'ththththtetewrtwerwer'
		if int(fourpin[x][fourpin[x].index('每箱规格')+1].split("CM")[1])>maxfour:
			maxfour=int(fourpin[x][fourpin[x].index('每箱规格')+1].split("CM")[1])

	print maxfour,'maxfour'

	maxthree=0
	for x in range(0,len(threepin)):
		print int(threepin[x][threepin[x].index('每箱规格')+1].split("CM")[1]),'ththththtetewrtwerwer'
		if int(threepin[x][threepin[x].index('每箱规格')+1].split("CM")[1])>maxthree:
			maxthree=int(threepin[x][threepin[x].index('每箱规格')+1].split("CM")[1])

	print maxthree,'maxthree'

	maxtwo=0
	for x in range(0,len(twopin)):
		print int(twopin[x][twopin[x].index('每箱规格')+1].split("CM")[1]),'twewewewewqrqfdqewfqaef'
		if int(twopin[x][twopin[x].index('每箱规格')+1].split("CM")[1])>maxtwo:
			maxtwo=int(twopin[x][twopin[x].index('每箱规格')+1].split("CM")[1])

	print maxtwo,'maxtwo'
	#none处理

	# for x in range (len(nonetype)):
	# 	lasttype.append(nonetype[x])
	for y in range (0,maxfive+1): 
		for x in range (len(fivepin)):
			if int(fivepin[x][fivepin[x].index('每箱规格')+1].split("CM")[1])==y:
				lasttype.append(fivepin[x])
	for y in range (0,maxfour+1): 
		for x in range (len(fourpin)):
			if int(fourpin[x][fourpin[x].index('每箱规格')+1].split("CM")[1])==y:
				lasttype.append(fourpin[x])
	for y in range (0,maxthree+1): 
		for x in range (len(threepin)):
			if int(threepin[x][threepin[x].index('每箱规格')+1].split("CM")[1])==y:
				lasttype.append(threepin[x])

	for y in range (0,maxtwo+1): 
		for x in range (len(twopin)):
			if int(twopin[x][twopin[x].index('每箱规格')+1].split("CM")[1])==y:
				lasttype.append(twopin[x])
	print lasttype,'last'

	
	for x in range (len(duzi)):
		lasttype.append(duzi[x])
	for x in range (len(zhuangwan)):
		lasttype.append(zhuangwan[x])
	for x in range (len(pi)):
		lasttype.append(pi[x])
	for x in range (len(chilun)):
		lasttype.append(chilun[x])


	return lasttype

#对价格进行处理
def chuchong(content):
	zxtype=[]
	zxzongtype=[]
	chongfu=[]
	filename(content)
	typee=[]
	a=0	
	#获取所有的sheet
	Sheetname=workbook.sheet_names()
	print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		print "第",name+1,"个sheet："
		#获取所有的行数
		nrows=table.nrows
		# if not nrows:
		# 	print "		空页"
		for n in range(nrows):
			#获取单行内容
			a=table.row_values(n)
			# print "		第",n+1,"行：",
			b=table.row_values(0)
			c=[]
			d=[]#新增的数组
			for l in range(len(a)): 
				c.append(b[l])
				c.append(a[l])

			print c
			print c[c.index('型号')+1],'saaaaaaaaaaaa'
			print type(c[c.index('型号')+1])
			if type(c[c.index('型号')+1])==type(u'asd'):
				if '-' in c[c.index('型号')+1]:
					try:
						c[c.index('型号')+1]=int(c[c.index('型号')+1].split("-")[1])
					except:
						c[c.index('型号')+1]=c[c.index('型号')+1].split("-")[1]		
				else:
					try:
						c[c.index('型号')+1]=int(c[c.index('型号')+1])
					except:
						print 'whith a'



					# print c[c.index('型号')+1].split("-")[1],type(c[c.index('型号')+1].split("-")[1])
					# if is_num(c[c.index('型号')+1].split("-")[1]):
					# 	c[c.index('型号')+1]=int(c[c.index('型号')+1].split("-")[1])
					# 	print 'sadasdasdas1111111'
					# c[c.index('型号')+1]=c[c.index('型号')+1].split("-")[1]
			zxtype.append(c[c.index('型号')+1])
			zxtype.append(c[c.index('客户编号')+1])
			jiage=c[c.index('单价')+1]
			print jiage,type(jiage)
			try:
				jiage=round(jiage,2)
			except:
				print 'head'
			print jiage
			zxtype.append(jiage)
			zxzongtype.append(zxtype)
			zxtype=[]
	zxzongtype1=[]
	for i in zxzongtype:
		if i not in zxzongtype1:
			zxzongtype1.append(i)

	print zxzongtype1
	flag=0
	for x in range (0,len(zxzongtype1)-2):
		if x>=len(zxzongtype1)-2 :
			break
		# zxtype.append(zxzongtype1[x])
		for y in range (x+1,len(zxzongtype1)-1):
			if y>=len(zxzongtype1)-1:
				break
			# print len(zxzongtype1),y,x
			if flag==1:#删除之后要倒退一位
				if zxzongtype1[y-1][2]=='':
					del zxzongtype1[y-1]
				elif  zxzongtype1[x][0]==zxzongtype1[y-1][0] and zxzongtype1[x][2]==zxzongtype1[y-1][2]:
					del zxzongtype1[y-1]
				else : flag=0
			if zxzongtype1[y][2]=='':
				del zxzongtype1[y]
			elif zxzongtype1[x][0]==zxzongtype1[y][0] and zxzongtype1[x][2]==zxzongtype1[y][2] and flag==0:	
				# print zxzongtype1[y]			
				del zxzongtype1[y]
				flag=1

	print zxzongtype1,'zxzongtype1'

	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	for x in range(len(zxzongtype1)):
		for y in range(0,3):
			if is_chinese(zxzongtype1[x][y]):
				zxzongtype1[x][y].encode('utf-8')
			# elif not four[i] nd four[i]!=0:
			# 	print "空值",
			elif is_num(zxzongtype1[x][y])==1:
				if math.modf(zxzongtype1[x][y])[0]==0 or zxzongtype1[x][y]==0:#获取数字的整数和小数
					zxzongtype1[x][y]=int(zxzongtype1[x][y])#将浮点数化成整数
			sheet1.write(x,y,zxzongtype1[x][y])		
	book.save('10.xls')#存储excel
	book = xlrd.open_workbook('10.xls')


 	flag=0
	for x in range (0,len(zxzongtype1)-2):
		if x>=len(zxzongtype1)-2 :
			break
		# zxtype.append(zxzongtype1[x])
		for y in range (x+1,len(zxzongtype1)-1):
			if y>=len(zxzongtype1)-1:
				break
			# print len(zxzongtype1),y,x
			if flag==1:#删除之后要倒退一位
				if  zxzongtype1[x][0]==zxzongtype1[y-1][0]:
					chongfu.append(zxzongtype1[y-1][0])
				else : flag=0
			if zxzongtype1[x][0]==zxzongtype1[y][0] and flag==0:	
				# print zxzongtype1[y]			
				chongfu.append(zxzongtype1[y][0])
				flag=1
	print chongfu,'chongfufufufu'



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
#最终的处理
def finaldeal(num,a,alltype):
	print a,'aaaaaaaaaaaaaa'
	less=a[0][num]
	buman=[]
	finalway=[]
	for l in range(len(a)):
		global wanquan
		global buman
		# if a[x][3]==0:
		# 	print'sbhengsda sdasasd'
		# 	wanquan.append(a[x])

		if a[l][num]<less:
			less=a[l][num] 

		if less==0 and a[l][num]==0:
			print'sbhengsda sdasasd'
			wanquan.append(a[l])
		elif a[l][num]!=0:
			buman.append(a[l])	
	print less,'!!!!!!!!!!!!!!'
	print wanquan
	print buman
	if wanquan!=[]:
		global finalway
		print wanquan
		less=wanquan[0][num+1]
		for y in range(len(wanquan)):
			global less
			if less>wanquan[y][num+1]:
				less=wanquan[y][num+1]
		print less,'wanquan'
		for z in range(len(wanquan)):
			if 	wanquan[z][num+1]==less:
				finalway= wanquan[z][0:num+1]
		print finalway,'finallway'#最终方案

	else: 
		global finalway
		alltype.sort()
		geshu=0
		lessbwuan=[]
		needdeal=[]
		for u in range(len(alltype)):
			if alltype[u]>buman[0][num]:
				lessnum=alltype[u]-buman[0][num]
				break
		for j in range(len(buman)):
			for m in range(len(alltype)):
				if alltype[m]>buman[j][num]:
					if alltype[m]-buman[j][num]<lessnum:
						lessnum=alltype[m]-buman[j][num]
						break
		print lessnum,'lessnum'
		for i in range(len(buman)):
			for j in range(len(alltype)):
				if alltype[j]>buman[i][num] and alltype[j]-buman[i][num]==lessnum:
					print alltype[j],buman[i][num],'gigiiggigigigigigi'
					geshu=geshu+1
					needdeal.append(buman[i])
					break

		less=0
		print needdeal,'needdeal'
		print geshu,'gegegegegegegeg'
		if len(needdeal)==1:
			finalway=needdeal[0][0:num+1]
			print finalway,'finaywwwywywyywywywywywywywy'
			
		else: 
			for k in range(len(alltype)):
				if needdeal[0][num]<alltype[k]  and needdeal[0][num]/alltype[k]<1:
					less=needdeal[0][num]/alltype[k]
					break


			print less,'lessssssssssssssssssssssssssssss'
			for i in range(len(needdeal)):
				for k in range(len(alltype)):
					if needdeal[i][num]<alltype[k] and needdeal[i][num]/alltype[k]<1:
						if needdeal[i][num]/alltype[k]>less:
							less=needdeal[i][num]/alltype[k]
						break
			print less,'lessssssssssssssssssssssssssssss'
			for i in range(len(needdeal)):
				for k in range(len(alltype)):
					if needdeal[i][num]<alltype[k] and needdeal[i][num]/alltype[k]<1:
						if needdeal[i][num]/alltype[k]==less:
							finalway=needdeal[i][0:num+1]
						break
			print finalway,'finallway'#最终方案
#最终处理后的输出
def finalwrite(finalway,alltype,gett):
	alltype.sort()
	alltype.reverse()
	for p in range(len(finalway)-1):
		for y in range(len(alltype)):
			if finalway[p]!=0:
				for z in range(len(gett)):
					for j in range(len(gett[z])):
						if gett[z][gett[z].index('型号')+1]==findell[0][0] and gett[z][gett[z].index('每箱规格')+1]=='none':

							global s1
							gettspec=gett[z][:]
							maxxxxx=gett[z][gett[z].index('需求数量')+1]
							gettspec[gettspec.index('数量')+1]=alltype[p]
							gettspec[gettspec.index('总件数')+1]=finalway[p]
							gettspec[gettspec.index('需求数量')+1]=finalway[p]*alltype[p]
							gettspec[gettspec.index('净重')+1]=gettspec[gettspec.index('需求数量')+1]/maxxxxx*gettspec[gettspec.index('净重')+1]
							gettspec[gettspec.index('毛重')+1]=gettspec[gettspec.index('需求数量')+1]/maxxxxx*gettspec[gettspec.index('毛重')+1]
							gettspec[gettspec.index('总重量')+1]=gettspec[gettspec.index('需求数量')+1]/maxxxxx*gettspec[gettspec.index('总重量')+1]
							for h in range(len(zxtype)):
								if  zxtype[h][1]==alltype[p] and zxtype[h][0]==gett[z][gett[z].index('内盒')+1]:
									# print gett[z][gett[z].index('型号')+1],findell[0][0]
									# print zxtype[h][1],alltype[p],'xaaaaaaaaaaaaaxaxaxaaaaaxa'
									gettspec[gettspec.index('每箱规格')+1]=zxtype[h][2]+'全'
									try:
										gettspec[gettspec.index('毛重')+1]=gettspec[gettspec.index('净重')+1]+zxtype[h][3]
										gettspec[gettspec.index('总重量')+1]=gettspec[gettspec.index('毛重')+1]
									except:
										print gettspec[gettspec.index('型号')+1]+'zhongliang is woring'
							s1=gettspec[:]
							
						
							if s1!=[] and s1 not in gett:
								print s1,'s333333333333333333333333'
								gett.append(s1)
	print finalway,'cc!!@!@!@!@cc!!@!@!@!@cc!!@!@!@!@cc!!@!@!@!@cc!!@!@!@!@'
	if finalway[-1]!=0:
		print 'succcccccccccccccccccccccccccccccc!!@!@!@!@'
		for z in range(len(gett)):
			for j in range(len(gett[z])):
				if gett[z][gett[z].index('型号')+1]==findell[0][0] and gett[z][gett[z].index('每箱规格')+1]=='none':
					global s1
					gettspec=gett[z][:]
					maxxxxx=gett[z][gett[z].index('需求数量')+1]
					gettspec[gettspec.index('总件数')+1]=1
					gettspec[gettspec.index('数量')+1]=finalway[-1]
					gettspec[gettspec.index('需求数量')+1]=finalway[-1]
					gettspec[gettspec.index('净重')+1]=gettspec[gettspec.index('需求数量')+1]/maxxxxx*gettspec[gettspec.index('净重')+1]
					gettspec[gettspec.index('毛重')+1]=gettspec[gettspec.index('需求数量')+1]/maxxxxx*gettspec[gettspec.index('毛重')+1]
					gettspec[gettspec.index('总重量')+1]=gettspec[gettspec.index('毛重')+1]
					alltype.sort()
					# print alltype
					flaaaa=0
									
					for j in range(len(alltype)):
						for h in range(len(zxtype)):	
							if alltype[j]>finalway[-1] and flaaaa==0:
								if zxtype[h][1]==alltype[j] and zxtype[h][0]==gett[z][gett[z].index('内盒')+1]:
									# print zxtype[h][2],zxtype[h][1],zxtype[h][0],alltype[j],gett[z][gett[z].index('内盒')+1]
									gettspec[gettspec.index('每箱规格')+1]=zxtype[h][2]+'最终'+str(zxtype[h][1])
									try:
										gettspec[gettspec.index('毛重')+1]=gettspec[gettspec.index('净重')+1]+zxtype[h][3]
										gettspec[gettspec.index('总重量')+1]=gettspec[gettspec.index('毛重')+1]
									except:
										print gettspec[gettspec.index('型号')+1]+'zhongliang is woring'
									flaaaa=1
									# print zxtype[h][2],'asdasdsadasdasdasdasdasdsa'
					s1=gettspec[:]
					# print s1,'whatwahtwahtawhttahwt'
					if s1!=[] and s1 not in gett:
						gett.append(s1)	

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

#带颜色的读取
def filename1(content):
	#打开文件
	global workbook,file_excel
	file_excel=str(content)
	file=(file_excel+'.xls').decode('utf-8')#文件名及中文合理性
	if not os.path.exists(file):#判断文件是否存在
		file=(file_excel+'.xlsx').decode('utf-8')
		if not os.path.exists(file):
			print "文件不存在"
	workbook = xlrd.open_workbook(file,formatting_info=1)
	print 'suicce'##1213
#拼箱模块
def pinxiang(gett):
	pinx=[]
	e=[]
	for x in range(len(gett)):
		if gett[x][gett[x].index('每箱规格')+1]=='none':
			e.append(gett[x][gett[x].index('型号')+1])
			e.append(gett[x][gett[x].index('内盒')+1])
			e.append(gett[x][gett[x].index('数量')+1])
			pinx.append(e)
			e=[]
	print pinx,'pdddddddddddddddddddddddddddddddddddddddddddddddddddddddd'	
	pinxtype=[]
	for x in range(len(pinx)):
		if pinx[x][1] not in pinxtype and (type(pinx[x][1])==type(1.0) or type(pinx[x][1])==type(1)):
			pinxtype.append(pinx[x][1])
	print pinxtype,'qqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqqq'
	global pinxdiff

	pinxdiff=[]
	for x in range(len(pinxtype)):
		qqq=[]
		for y in range(len(pinx)):
			# print str(pinxtype[x]),str(pinx[y][1]),'?:?:?::'
			if pinx[y][1]==pinxtype[x]:
				qqq.append(pinx[y])
			elif str(int(pinxtype[x])) in  str(pinx[y][1]):#齿轮头处理方法
				a=pinx[y][1].split("*")[0]
				b=pinx[y][1].split("*")[1]
				pinx[y][1]=float(a)
				pinx[y][2]=math.ceil(pinx[y][2]/int(b))

				qqq.append(pinx[y])


		pinxdiff.append(qqq)
	print pinxdiff,'fdifidfidifidifdifi'

	print zxtype
	xiangzitype=[]
	for x in range (len(pinxtype)):
		aaa=[]
		for y in range(len(zxtype)):
			if pinxtype[x]==zxtype[y][0]:
				aaa.append(zxtype[y])
		xiangzitype.append(aaa)
	print xiangzitype,'xiangzitype'


	

	b=0
	for x in range (len(xiangzitype)):
		for y in range (len(xiangzitype[x])):
			for z in range (len(pinxdiff)):
				for j in range (0,len(pinxdiff[z])-4):
					for k in range (1,len(pinxdiff[z])-3):
						for i in range (2,len(pinxdiff[z])-2):
							for l in range(3,len(pinxdiff[z])-1):
								for m in range(4,len(pinxdiff[z])):
									if len(pinxdiff[z][j])==3 and len(pinxdiff[z][k])==3 and len(pinxdiff[z][i])==3 and len(pinxdiff[z][l])==3 and len(pinxdiff[z][m])==3:
										if xiangzitype[x][y][0]==pinxdiff[z][j][1]==pinxdiff[z][k][1]==pinxdiff[z][i][1]==pinxdiff[z][l][1]==pinxdiff[z][m][1]:
											if pinxdiff[z][j][2]+pinxdiff[z][k][2]+pinxdiff[z][i][2]+pinxdiff[z][l][2]+pinxdiff[z][m][2]==xiangzitype[x][y][1]:
												if pinxdiff[z][j][0]!=pinxdiff[z][k][0] and pinxdiff[z][j][0]!=pinxdiff[z][i][0] and pinxdiff[z][j][0]!=pinxdiff[z][l][0] and pinxdiff[z][k][0]!=pinxdiff[z][i][0] and pinxdiff[z][k][0]!=pinxdiff[z][l][0] and pinxdiff[z][i][0]!=pinxdiff[z][l][0] and pinxdiff[z][j][0]!=pinxdiff[z][m][0] and pinxdiff[z][k][0]!=pinxdiff[z][m][0] and pinxdiff[z][i][0]!=pinxdiff[z][m][0] and pinxdiff[z][l][0]!=pinxdiff[z][m][0]:
													a='5拼'+xiangzitype[x][y][2]+str(b)
													c=xiangzitype[x][y][3]
													pinxdiff[z][j].append(a)
													pinxdiff[z][k].append(a)
													pinxdiff[z][i].append(a)
													pinxdiff[z][l].append(a)
													pinxdiff[z][m].append(a)
													pinxdiff[z][j].append(c)
													pinxdiff[z][k].append(c)
													pinxdiff[z][i].append(c)
													pinxdiff[z][l].append(c)
													pinxdiff[z][m].append(c)
													b=b+1
													print pinxdiff[z][j],pinxdiff[z][k],pinxdiff[z][i],pinxdiff[z][l],pinxdiff[z][m],'five'


	b=0
	for x in range (len(xiangzitype)):
		for y in range (len(xiangzitype[x])):
			for z in range (len(pinxdiff)):
				for j in range (0,len(pinxdiff[z])-3):
					for k in range (1,len(pinxdiff[z])-2):
						for i in range (2,len(pinxdiff[z])-1):
							for l in range(3,len(pinxdiff[z])):
								if len(pinxdiff[z][j])==3 and len(pinxdiff[z][k])==3 and len(pinxdiff[z][i])==3 and len(pinxdiff[z][l])==3:
									if xiangzitype[x][y][0]==pinxdiff[z][j][1]==pinxdiff[z][k][1]==pinxdiff[z][i][1]==pinxdiff[z][l][1]:
										if pinxdiff[z][j][2]+pinxdiff[z][k][2]+pinxdiff[z][i][2]+pinxdiff[z][l][2]==xiangzitype[x][y][1]:
											if pinxdiff[z][j][0]!=pinxdiff[z][k][0] and pinxdiff[z][j][0]!=pinxdiff[z][i][0] and pinxdiff[z][j][0]!=pinxdiff[z][l][0] and pinxdiff[z][k][0]!=pinxdiff[z][i][0] and pinxdiff[z][k][0]!=pinxdiff[z][l][0] and pinxdiff[z][i][0]!=pinxdiff[z][l][0]:
												a='4拼'+xiangzitype[x][y][2]+str(b)
												c=xiangzitype[x][y][3]
												pinxdiff[z][j].append(a)
												pinxdiff[z][k].append(a)
												pinxdiff[z][i].append(a)
												pinxdiff[z][l].append(a)
												pinxdiff[z][j].append(c)
												pinxdiff[z][k].append(c)
												pinxdiff[z][i].append(c)
												pinxdiff[z][l].append(c)
												b=b+1
												print pinxdiff[z][j],pinxdiff[z][k],pinxdiff[z][i],pinxdiff[z][l],'four'







	b=0
	for x in range (len(xiangzitype)):
		for y in range (len(xiangzitype[x])):
			for z in range (len(pinxdiff)):
				for j in range (0,len(pinxdiff[z])-2):
					for k in range (1,len(pinxdiff[z])-1):
						for i in range (2,len(pinxdiff[z])):
							if len(pinxdiff[z][j])==3 and len(pinxdiff[z][k])==3 and len(pinxdiff[z][i])==3:
								if xiangzitype[x][y][0]==pinxdiff[z][j][1]==pinxdiff[z][k][1]==pinxdiff[z][i][1]:
									if pinxdiff[z][j][2]+pinxdiff[z][k][2]+pinxdiff[z][i][2]==xiangzitype[x][y][1]:
										if pinxdiff[z][j][0]!=pinxdiff[z][k][0] and pinxdiff[z][j][0]!=pinxdiff[z][i][0] and pinxdiff[z][k][0]!=pinxdiff[z][i][0]:
											a='3拼'+xiangzitype[x][y][2]+str(b)
											c=xiangzitype[x][y][3]
											pinxdiff[z][j].append(a)
											pinxdiff[z][k].append(a)
											pinxdiff[z][i].append(a)
											pinxdiff[z][j].append(c)
											pinxdiff[z][k].append(c)
											pinxdiff[z][i].append(c)
											b=b+1
											print pinxdiff[z][j],pinxdiff[z][k],pinxdiff[z][i],'threepinnnnnnn'






	b=0
	for x in range (len(xiangzitype)):
		for y in range (len(xiangzitype[x])):
			for z in range (len(pinxdiff)):
				for j in range (0,len(pinxdiff[z])-1):
					for k in range (1,len(pinxdiff[z])):
						if len(pinxdiff[z][j])==3 and len(pinxdiff[z][k])==3:
							# print '1111111111111111111111111'
							if xiangzitype[x][y][0]==pinxdiff[z][j][1]==pinxdiff[z][k][1]:

								if pinxdiff[z][j][2]+pinxdiff[z][k][2]==xiangzitype[x][y][1]:
									if pinxdiff[z][j][0]!=pinxdiff[z][k][0]:

										a='2拼'+xiangzitype[x][y][2]+str(b)
										c=xiangzitype[x][y][3]
										pinxdiff[z][j].append(a)
										pinxdiff[z][k].append(a)
										pinxdiff[z][j].append(c)
										pinxdiff[z][k].append(c)
										b=b+1
	print pinxdiff,'!@!@!@!@!@@@@@@@@@@EEEEEEE'





	# pinxtyne=[]
	# zhongzhuan=[]
	# for y in pinxtype:
	# 	a=0
	# 	for x in range(len(pinx)):
	# 		if y ==pinx[x][1]:
	# 			a=a+pinx[x][2]
	# 	zhongzhuan.append(y)
	# 	zhongzhuan.append(a)
	# 	pinxtyne.append(zhongzhuan)
	# 	zhongzhuan=[]
	
	# print zxtype
	# mtype=[]
	# for y in range(len(pinxtyne)):
		
	# 	for x in range(len(zxtype)):
	# 		if pinxtyne[y][0]==zxtype[x][0]  and pinxtyne[y][1] ==zxtype[x][1] :
	# 			print "1231232131231232131"
	# 			pinxtyne[y].append(zxtype[x][2])

	# print pinx,'2222222222222222222222222'
	# for x in range (len(pinx)):
	# 	for y in range(len(pinxtyne)):
	# 		if len(pinxtyne[y])>2:
	# 			if pinx[x][1]==pinxtyne[y][0]:
	# 				a=pinxtyne[y][2]+'拼箱:'+str(y)+'号'
	# 				pinx[x].append(a)
	# print pinx,'fangfanei de pinx'





#初步处理后的输出
def firstout(needdiff,finalway,alltype,gett):
	namestr=['单向器','齿轮头','皮带轮','弹簧装置皮带轮','单向器散装规格']
	alltype.sort()
	alltype.reverse()
	for x in range(len(finalway)-1):
		for y in range(len(alltype)):
			if finalway[x]!=0 :
				filename('zx')
				Sheetname=workbook.sheet_names()
				for name in range(len(Sheetname)):
					table = workbook.sheets()[name]
					s1=[]
					global s2
					s2=[]
					nrows=table.nrows
					# print nrows,'qwioweqfuqoifhoiqhfoiqfhi'
					for n in range(nrows):
						a=table.row_values(n)
						b=table.row_values(0)
						c=[]
						d=[]#新增的数组
						for l in range(len(a)): 
							c.append(b[l])
							c.append(a[l])
							e=len(c)
							for d in range(len(c)):
								# print x,len(c)
								if c[d]==' ' and c[d+1]==' ':
									c=c[0:d]
									break

						c.append(unicode('品名','utf-8'))
						c.append(unicode(namestr[name],'utf-8'))#为了区分那个sheet
						c.append(unicode('单位','utf-8'))
						c.append(unicode('只','utf-8'))
						if is_num(c[c.index('型号')+1])==1:
							if math.modf(c[c.index('型号')+1])[0]==0 or c[c.index('型号')+1]==0:#获取数字的整数和小数
								c[c.index('型号')+1]=int(c[c.index('型号')+1])#将浮点数化成整数

						# print needdiff[0][0],alltype[x],alltype[x],c[c.index('型号')+1],c[c.index('数量')+1]
						if c[c.index('型号')+1]==needdiff[0][0] and c[c.index('数量')+1]==alltype[x]:
							print 'succcccccccccccccccccccccccccccccc???????'
							global s1
							c.append(unicode('总件数','utf-8'))
							c.append(finalway[x])
							c.append(unicode('需求数量','utf-8'))
							c.append(finalway[x]*int(c[c.index('数量')+1]))
							c.append(unicode('总重量','utf-8'))
							c.append(finalway[x]*c[c.index('毛重')+1])
							if 'CM' not in c[c.index('每箱规格')+1]:
								c[c.index('每箱规格')+1]=c[c.index('每箱规格')+1]+'CM'
							s1=c[:]
							print s1,'s1111111111111111111111111'
					print s1,'s2222222222222222222222222222222222222222222'
					if s1!=[] and s1 not in gett:
						print s1,'s1111111111111111111111111'
						gett.append(s1)
	if finalway[-1]!=0:
		filename('zx')
		Sheetname=workbook.sheet_names()
		for name in range(len(Sheetname)):
			table = workbook.sheets()[name]
			global s1
			s1=[]
			global s2
			s2=[]
			nrows=table.nrows
			# print nrows,'qwioweqfuqoifhoiqhfoiqfhi'
			for n in range(nrows):
				a=table.row_values(n)
				b=table.row_values(0)
				c=[]
				d=[]#新增的数组
				for l in range(len(a)): 
					c.append(b[l])
					c.append(a[l])
					e=len(c)
					for d in range(len(c)):
						# print x,len(c)
						if c[d]==' ' and c[d+1]==' ':
							c=c[0:d]
							break

				c.append(unicode('品名','utf-8'))
				c.append(unicode(namestr[name],'utf-8'))#为了区分那个sheet
				c.append(unicode('单位','utf-8'))
				c.append(unicode('只','utf-8'))
				if c[c.index('型号')+1]==needdiff[0][0] and c[c.index('数量')+1]==alltype[1]:
					print 'succcccccccccccccccccccccccccccccc'
					global s1
					c.append(unicode('总件数','utf-8'))
					c.append(1)
					dd=c[c.index('毛重')+1]-c[c.index('净重')+1]
					c[c.index('净重')+1]=c[c.index('净重')+1]*(finalway[-1]/c[c.index('数量')+1])
					c[c.index('毛重')+1]=c[c.index('净重')+1]+dd
					c[c.index('数量')+1]=finalway[-1]
					c.append(unicode('需求数量','utf-8'))
					c.append(finalway[-1])
					c.append(unicode('总重量','utf-8'))
					c.append(finalway[-1]/(c[c.index('数量')+1])*(c[c.index('毛重')+1]))
					c[c.index('每箱规格')+1]='none'
					s1=c[:]
					print s1,'s1111111111111111111111111'
			print s1,'s1111111111111111111111111'
			if s1!=[] and s1 not in gett:
				print s1,'s1111111111111111111111111'
				gett.append(s1)	


# def filename():
# 	#打开文件
# 	global workbook,file_excel
# 	content="请输入要打开的excel文件名（不带后缀）:"
# 	while True:
# 		file_excel=raw_input(content)
# 		print file_excel
# 		file=(file_excel+'.xls').decode('utf-8')#文件名及中文合理性
# 		if not os.path.exists(file):#判断文件是否存在
# 			file=(file_excel+'.xlsx').decode('utf-8')
# 			if not os.path.exists(file):
# 				print "文件不存在"
# 				continue
# 		workbook = xlrd.open_workbook(file)
# 		break

def check():
	last=[]
	lastitem=[]
	filename(5)
	Sheetname=workbook.sheet_names()
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		#获取所有的行数
		nrows=table.nrows
		for n in range(nrows):
			#获取单行内容
			a=table.row_values(n)
			a=[i for i in a if i]#清除列表中后面的空值
			for i in range(len(a)):			
				if is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				last.append(a[0])
	# print last,'neeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeed'

	queshi=[]
	for x in range(len(needitem)):
		if needitem[x] not in last:
			queshi.append(needitem[x])
	print queshi,'lake type'


	










def readexcel(content):
	filename(content)
		
	#获取所有的sheet
	Sheetname=workbook.sheet_names()
	print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		print "第",name+1,"个sheet："
		#获取所有的行数
		nrows=table.nrows
		if not nrows:
			print "		空页"
		for n in range(nrows):
			#获取单行内容
			a=table.row_values(n)
			print "		第",n+1,"行：",
			a=[i for i in a if i]#清除列表中后面的空值
			if not a:
				print "空行",
			for i in range(len(a)):			
				if is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数
				if type(a[i])==type(u'asd'):
					if '-' in a[i]:
						try:
							a[i]=int(a[i].split("-")[1])
						except:
							a[i]=a[i].split("-")[1]	
					try:
						if ' ' in a[i]:
							try:
								a[i]=int(a[i].split(" ")[1])
							except:
								a[i]=a[i].split(" ")[1]
					except:
						print 'isint?' 		
				need.append(a[i])


			# if not a:
			# 	print "空行",
			# for i in range(len(a)):				
			# 	if is_chinese(a[i]):
			# 		print a[i].encode('utf-8' ),'  ',
			# 		need.append(a[i])
			# 	elif not a[i] and a[i]!=0:
			# 		print "空值",
			# 	elif is_num(a[i])==1:
			# 		if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
			# 			a[i]=int(a[i])#将浮点数化成整数
			# 			need.append(a[i])
			# 		print a[i],'   ',
			# 	else:
			# 		print a[i],'  ',
			# print ''

	print need,'neeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeed'
	for x in range(len(need)):
		if (x%2 ==0):
			needitem.append(need[x])


	#处理重复的数据
	c=list(set(needitem))
	global needitem
	needitem=c
	print needitem,'nedfifjiosadjfoiasfoiasiaof'
	for x in range (len(needitem)):
		v=0
		he=0
		for y in range( len(need)):
			if need[y]==needitem[x]:
				v=v+1
		if v>1:
			print 'vbvbvbvbvbvbvbvb',v,needitem[x]
			flag=0
			for z in range (len(need)):
				if z >=len(need):
					break
				if flag==1:
					if need[z-1]==needitem[x]:
						he=he+need[z]
						del need[z-1]
						del need[z-1]
					else : flag=0
				if need[z]==needitem[x] and flag==0:
					he=he+need[z+1]
					del need[z]
					del need[z]
					flag=1
					print need,'asdwqewqfaqvcaqfqf',he
		if he!=0:
			print need,'nenenenenendndnendndnenndn',he ,needitem[x]
			need.append(needitem[x])
			need.append(he)
	print need,'nenenenenendndnendndnenndn',he


def readexcel2(content):
	zong=0
	sheng=0
	zongliang=0
	flag=0

	cunzai=[]#判断内容是否存在
	yuanshuliang=0#原先的数量
	yuanxianzong=0#原先总需求
	zxzongtype=[]
	global zxtype
	# pinx=[]#所有需要拼箱的产品
	print need
	for  x in range(len(needitem)):
		cunzai.append('0')
	print cunzai


	namestr=['单向器','齿轮头','皮带轮','弹簧装置皮带轮','单向器散装规格']
	for k in range(len(needitem)):#查找装箱里面所有和1相符的内容并将对应的需求数量加上去
		filename1(content)
		Sheetname=workbook.sheet_names()
		needdiff=[]
		for name in range(len(Sheetname)):
			table = workbook.sheets()[name]
			nrows=table.nrows
			for n in range(nrows):
				a=table.row_values(n)
				b=table.row_values(0)
				c=[]
				d=[]#新增的数组
				for l in range(len(a)): 
					c.append(b[l])
					c.append(a[l])
				c.append(unicode('品名','utf-8'))
				c.append(unicode(namestr[name],'utf-8'))#为了区分那个sheet
				c.append(unicode('单位','utf-8'))
				c.append(unicode('只','utf-8'))
				sheet = workbook.sheet_by_index(name)
				xfx = sheet.cell_xf_index(n, 0)
				xf = workbook.xf_list[xfx]
				bgx = xf.background.pattern_colour_index
				c.append(bgx)
			
				# zxtype.append(c[c.index('内盒')+1])
				# zxtype.append(c[c.index('数量')+1])
				# zxtype.append(c[c.index('每箱规格')+1])
				# zxzongtype.append(zxtype)
				# zxtype=[]
				# print needitem[k]

				# if needitem[k] in c and c[-1]==11:  #判断颜色
				if needitem[k] in c :
					if c[-1]==64 :
						for x in range(len(zxtype)):
							if c[c.index('内盒')+1]==zxtype[x][0] and c[c.index('每箱规格')+1]==zxtype[x][2]:
								if c[c.index('数量')+1]!=zxtype[x][1]:
									c[c.index('数量')+1]=zxtype[x][1]


					a=[]
					a.append(needitem[k])
					a.append(c[c.index('数量')+1])
					needdiff.append(a)

		print needdiff,'bnbnbnbbnbnbnbnbnbnbnbnbnbnbn'
#最初处理单匹配
		if len(needdiff)>0 :
			if len(needdiff)==1:
				for x in range(len(need)):
					if needitem[k]==need[x] :
				
						filename(content)
						Sheetname=workbook.sheet_names()
						for name in range(len(Sheetname)):
							table = workbook.sheets()[name]
							global s1
							s1=[]
							global s2
							s2=[]
							nrows=table.nrows
							# print nrows,'qwioweqfuqoifhoiqhfoiqfhi'
							for n in range(nrows):
								a=table.row_values(n)
								b=table.row_values(0)
								c=[]
								d=[]#新增的数组
								for l in range(len(a)): 
									c.append(b[l])
									c.append(a[l])
									e=len(c)
									for d in range(len(c)):
										# print x,len(c)
										if c[d]==' ' and c[d+1]==' ':
											c=c[0:d]
											break

								c.append(unicode('品名','utf-8'))
								c.append(unicode(namestr[name],'utf-8'))#为了区分那个sheet
								c.append(unicode('单位','utf-8'))
								c.append(unicode('只','utf-8'))
							# print c,'ccccccc    wath!!!!',
								# print needdiff[0][0],c[c.index('型号')+1],"xxxxxxxxxxxx!!!!!"
								if c[c.index('型号')+1]==needdiff[0][0] and needdiff[0][1]!='':
									print 'succcccccccccccccccccccccccccccccc'
									if need[x+1]%c[c.index('数量')+1] !=0 and int(need[x+1]/c[c.index('数量')+1])>0:#多出
										print '111111111111111111111111111111111111111111111'
										d=c[:]
										flag=1
										yuanxianzong=need[x+1]
										zong=int(need[x+1]/c[c.index('数量')+1])
										c.append(unicode('总件数','utf-8'))
										c.append(zong)
										zongliang=zong*c[c.index('数量')+1]
										print '总量',zongliang,'袁总量',yuanxianzong
										c.append(unicode('需求数量','utf-8'))
										c.append(zongliang)
										c.append(unicode('总重量','utf-8'))
										if 'CM' not in c[c.index('每箱规格')+1]:
											c[c.index('每箱规格')+1]=c[c.index('每箱规格')+1]+'CM'

										try:
											c.append(zong*c[c.index('毛重')+1])
										except:
											c.append(0)
											print 'que'
										s1=c[:]
										d[d.index('数量')+1]=yuanxianzong-zongliang
										d.append(unicode('总件数','utf-8'))
										d.append(1)
										d.append(unicode('需求数量','utf-8'))
										d.append(yuanxianzong-zongliang)
										try:
											d[d.index('净重')+1]=d[d.index('净重')+1]*(d[d.index('数量')+1]/c[c.index('数量')+1])
										except:
											d[d.index('净重')+1]=0
											print 'que'
										try:
											d[d.index('毛重')+1]=d[d.index('净重')+1]+((c[c.index('毛重')+1])-(c[c.index('净重')+1]))
										except:
											d[d.index('毛重')+1]=0
											print 'que'
										d[d.index('每箱规格')+1]='none'
										d.append(unicode('总重量','utf-8'))
										d.append((d[d.index('毛重')+1]))
										s2=d[:]
										print s1,s2
									elif int(need[x+1]/c[c.index('数量')+1])==0:#一箱都不满
											print '12222222222222222222222222222222222222222222222'
											yuanshuliang=c[c.index('数量')+1]
											c[c.index('数量')+1]=need[x+1]
											c.append(unicode('总件数','utf-8'))
											c.append(need[x+1]/c[c.index('数量')+1])
											c.append(unicode('需求数量','utf-8'))
											c.append(need[x+1])
											try:
												c[c.index('净重')+1]=c[c.index('净重')+1]*(c[c.index('数量')+1]/yuanshuliang)
											except:
												c[c.index('净重')+1]=0
												print 'que'
											try:
												c[c.index('毛重')+1]=c[c.index('毛重')+1]*(c[c.index('数量')+1]/yuanshuliang)
											except:
												c[c.index('毛重')+1]=0
												print 'que'	

											c.append(unicode('总重量','utf-8'))
											try:
												c.append(need[x+1]/c[c.index('数量')+1]*c[c.index('毛重')+1])
											except:
												c.append(0)

											c[c.index('每箱规格')+1]='none'
											s1=c[:]
									else: #正常
										print '3333333333333333333333333333333333333333'
										# print c
										c.append(unicode('总件数','utf-8'))
										c.append(need[x+1]/c[c.index('数量')+1])
										c.append(unicode('需求数量','utf-8'))
										c.append(need[x+1])
										c.append(unicode('总重量','utf-8'))
										try:
											c.append(need[x+1]/c[c.index('数量')+1]*c[c.index('毛重')+1])
										except:
											c.append(0)
											print 'que'
										if 'CM' not in c[c.index('每箱规格')+1]:
											c[c.index('每箱规格')+1]=c[c.index('每箱规格')+1]+'CM'
										s1=c[:]
										# print s1
								if c[c.index('型号')+1]==needdiff[0][0] and needdiff[0][1]=='':
									c.append(unicode('总件数','utf-8'))
									c.append('')
									c.append(unicode('需求数量','utf-8'))
									c.append(need[x+1])
									c.append(unicode('总重量','utf-8'))
									c.append('')
									s1=c[:]
									print c[c.index('型号')+1],'xxxxxxxxxhxhxhhxhxhh'



							if s1!=[]:
								gett.append(s1)
								if flag==1:
									gett.append(s2)
									flag=0



	#最初处理二匹配
			elif len(needdiff)==2:
				a=[]
				b=[]
				wanquan=[]#全装箱
				alltype=[]
				print need
				print needdiff
				for k in range(len(needdiff)):
						alltype.append(needdiff[k][1])
						alltype.sort()
						alltype.reverse()
				print alltype,'typetype'
				for l in range(len(need)):
					for k in range(len(needdiff)):
						if need[l]==needdiff[k][0]:
							zongshu=need[l+1]
				print zongshu,'zonsgaid'
				print zongshu//alltype[0]
				for x in range(int(zongshu//alltype[0])+1):
					for y in range(int(zongshu//alltype[1])+1):
						if alltype[0]*x+alltype[1]*y<=zongshu and zongshu-alltype[0]*x-alltype[1]*y<alltype[0]:
							global a
							b=[]
							b.append(x)
							b.append(y)
							b.append(zongshu-alltype[0]*x-alltype[1]*y)
							b.append(x+y)
						if b!=[] and b not in a:
				
						# print 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
							a.append(b)
							print alltype[0]*x+alltype[1]*y-zongshu
				print a,'aaaaaaaaaaaaaa'
				less=a[0][2]
				buman=[]
				finalway=[]
				for x in range(len(a)):
					global wanquan
					global buman
					# if a[x][3]==0:
					# 	print'sbhengsda sdasasd'
					# 	wanquan.append(a[x])

					if a[x][2]<less:
						less=a[x][2] 

					if less==0 and a[x][2]==0:
						print'sbhengsda sdasasd'
						wanquan.append(a[x])
					elif a[x][2]!=0:
						buman.append(a[x])	
				print less,'!!!!!!!!!!!!!!'
				print wanquan
				print buman
				if wanquan!=[]:
					global finalway
					print wanquan
					less=wanquan[0][3]
					for y in range(len(wanquan)):
						global less
						print less,'wanquan'
						print wanquan[y][3]
						if less>wanquan[y][3]:
							less=wanquan[y][3]
					print less,'wanquan'
					for z in range(len(wanquan)):
						if 	wanquan[z][3]==less:
							finalway.append(wanquan[z][0:3])

				else: 
					global finalway
					lessbwuan=[]
					alltype.sort()
					flagg=0
					print alltype,'allllllllllllllllllllllllttttttttttttyyyyyyyyyyppppppppee'
					for k in range(len(alltype)):
						if buman[0][2]<alltype[k] and flagg==0 and buman[0][2]/alltype[k]<1:
							less=buman[0][2]/alltype[k]
							flagg=1
							

					for y in range(len(buman)):
						flagg=0
						print '??!?!?!??!!??!?!!?!?!??'
						for k in range(len(alltype)):
							if buman[y][2]<alltype[k] and flagg==0:
								global lessbwuan
								global less
								print less,'buwan！！！！！！！！！！！！！！！！！！'
								print alltype[k]
								if less<=buman[y][2]/alltype[k]:
									print 'wath111i1hio1hohoh'
									less=buman[y][2]/alltype[k]
									lessbwuan.append(less)
									lessbwuan.append(alltype[k])
									flagg=1
					print less,'buwan'
					print buman,'buwnbuwanbuwan'

					for z in range(len(buman)):
						global lessbwuan
						print lessbwuan
						print buman[z][2],len(buman),z,lessbwuan[0]*lessbwuan[1]
						if 	buman[z][2]==lessbwuan[0]*lessbwuan[1]:
							finalway.append(buman[z][0:3])
							print finalway,'differentway'
				print finalway,'minxiangfinalwua'
				minxiang=finalway[0][0]+finalway[0][1]
				for g in range(len(finalway)):
				 	minn=finalway[g][0]+finalway[g][1]
				 	if minn<=minxiang:
				 		minxiang=minn
				 		print minxiang,'mingxiang'
				for j in range(len(finalway)):
					if finalway[j][0]+finalway[j][1]==minxiang:
						global finalway

						finalway=finalway[j]
						print finalway,'fffffffffffffffffffffff'
						break
				print finalway,'finallway'#最终方案

				print needdiff,'neeeeedddd'
				print alltype
				firstout(needdiff,finalway,alltype,gett)
			



	#最初处理三匹配
			elif len(needdiff)==3:
				a=[]
				b=[]
				wanquan=[]#全装箱
				alltype=[]
				print need
				print needdiff
				for k in range(len(needdiff)):
						alltype.append(needdiff[k][1])
						alltype.sort()
						alltype.reverse()
				print alltype,'typetype'
				for l in range(len(need)):
					for k in range(len(needdiff)):
						if need[l]==needdiff[k][0]:
							zongshu=need[l+1]
				print zongshu,'zonsgaid'
				print zongshu//alltype[0]
				for x in range(int(zongshu//alltype[0])+1):
					for y in range(int(zongshu//alltype[1])+1):
						for z in range(int(zongshu//alltype[2])+1):
							if alltype[0]*x+alltype[1]*y+alltype[2]*z<=zongshu and zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z<alltype[0]:
								global a
								b=[]
								b.append(x)
								b.append(y)
								b.append(z)
								b.append(zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z)
								b.append(x+y+z)
							if b!=[] and b not in a:
					
							# print 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
								a.append(b)
								print alltype[0]*x+alltype[1]*y+alltype[2]*z-zongshu
				print a,'aaaaaaaaaaaaaa'
				less=a[0][3]
				buman=[]
				finalway=[]
				for x in range(len(a)):
					global wanquan
					global buman
					# if a[x][3]==0:
					# 	print'sbhengsda sdasasd'
					# 	wanquan.append(a[x])

					if a[x][3]<less:
						less=a[x][3] 

					if less==0 and a[x][3]==0:
						print'sbhengsda sdasasd'
						wanquan.append(a[x])
					elif a[x][3]!=0:
						buman.append(a[x])	
				print less,'!!!!!!!!!!!!!!'
				print wanquan
				print buman
				if wanquan!=[]:
					global finalway
					print wanquan
					less=wanquan[0][4]
					for y in range(len(wanquan)):
						global less
						if less>wanquan[y][4]:
							less=wanquan[y][4]
					print less,'wanquan'
					for z in range(len(wanquan)):
						if 	wanquan[z][4]==less:
							finalway.append(wanquan[z][0:4])
							print finalway,'finallway'#最终方案

				else: 
					global finalway
					lessbwuan=[]
					alltype.sort()
					flagg=0
					print alltype,'allllllllllllllllllllllllttttttttttttyyyyyyyyyyppppppppee'
					for k in range(len(alltype)):
						if buman[0][3]<alltype[k] and flagg==0 and buman[0][3]/alltype[k]<1:
							less=buman[0][3]/alltype[k]
							flagg=1
							

					for y in range(len(buman)):
						flagg=0
						print '??!?!?!??!!??!?!!?!?!??'
						for k in range(len(alltype)):
							if buman[y][3]<alltype[k] and flagg==0:
								global lessbwuan
								global less
								print less,'buwan！！！！！！！！！！！！！！！！！！'
								print alltype[k]
								if less<=buman[y][3]/alltype[k]:
									print 'wath111i1hio1hohoh'
									less=buman[y][3]/alltype[k]
									lessbwuan.append(less)
									lessbwuan.append(alltype[k])
									flagg=1
					print less,'buwan'
					print buman,'buwnbuwanbuwan'

					for z in range(len(buman)):
						global lessbwuan
						print lessbwuan
						print buman[z][3],len(buman),z,lessbwuan[0]*lessbwuan[1]
						if 	buman[z][3]==lessbwuan[0]*lessbwuan[1]:
							finalway.append(buman[z][0:4])
							print finalway,'differentway'
					print finalway,'minxiangfinalwua'
				minxiang=finalway[0][0]+finalway[0][1]+finalway[0][2]
				for g in range(len(finalway)):
				 	minn=finalway[g][0]+finalway[g][1]+finalway[g][2]
				 	if minn<=minxiang:
				 		minxiang=minn
				 		print minxiang,'mingxiang'
				for j in range(len(finalway)):
					if finalway[j][0]+finalway[j][1]+finalway[j][2]==minxiang:
						global finalway

						finalway=finalway[j]
						print finalway,'fffffffffffffffffffffff'
						break
				print finalway,'finallway'#最终方案
				firstout(needdiff,finalway,alltype,gett)





	
	qwww=[]

	for x in range(len(needitem)):
		b=0
		for y in range(len(gett)):
			if gett[y][gett[y].index('型号')+1]==needitem[x]:
				b=1
				break
		if b==0:
			qwww.append(needitem[x])
	print qwww,'qwwwwwwwwwwwwwwwwwwwwwww'
	# 不存在
	if len(qwww)>0:
		for x in range (len(qwww)):
			for y in range (len(need)):
				if need[y]==qwww[x]:
					ne=[]
					ne.append(unicode('型号','utf-8'))
					ne.append(qwww[x])
					ne.append(unicode('只重量','utf-8'))
					ne.append('')
					ne.append(unicode('内盒','utf-8'))
					ne.append('')
					ne.append(unicode('数量','utf-8'))
					ne.append('')
					ne.append(unicode('净重','utf-8'))
					ne.append('')
					ne.append(unicode('毛量','utf-8'))
					ne.append('')
					ne.append(unicode('体积','utf-8'))
					ne.append('')
					ne.append(unicode('每箱规格','utf-8'))
					ne.append('')
					ne.append(unicode('重量','utf-8'))
					ne.append('')
					ne.append(unicode('品名','utf-8'))
					ne.append('')
					ne.append(unicode('单位','utf-8'))
					ne.append(unicode('只','utf-8'))
					ne.append(unicode('总件数','utf-8'))
					ne.append('')
					ne.append(unicode('需求数量','utf-8'))
					ne.append(need[y+1])
					ne.append(unicode('总重量','utf-8'))
					ne.append('')
					gett.append(ne)
	print gett,'getrtttttttttttttttttttttttttttttttttttttttttttttttt'

	pinx=[]
	pinxiang(gett)
	global pinxdiff
	print pinxdiff,'pinxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxpinxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'


	for x in range(len(pinxdiff)):
		for y in range(len(pinxdiff[x])):
			for z in range(len(gett)):
				if len(pinxdiff[x][y])>3 and pinxdiff[x][y][0]==gett[z][gett[z].index('型号')+1] and gett[z][gett[z].index('每箱规格')+1]=='none':
					gett[z][gett[z].index('每箱规格')+1]=pinxdiff[x][y][3]
					gett[z][gett[z].index('毛重')+1]='+'+str(pinxdiff[x][y][4])



	# for x in range(len(pinx)):
	# 	for y in range(len(gett)):
	# 			if pinx[x][0]==gett[y][gett[y].index('型号')+1] and gett[y][gett[y].index('每箱规格')+1]=='none':
	# 				print len(pinx),x
	# 				print len(gett), y
	# 				print len(gett[y])
	# 				print gett[y].index('每箱规格')
	# 				if len(pinx[x])>3:
	# 					gett[y][gett[y].index('每箱规格')+1]=pinx[x][3]


	print gett,'getrtttttttttttttttttttttttttttttttttttttttttttttttt'


	findell=[]
	for s in range(len(gett)):
		if gett[s][gett[s].index('每箱规格')+1]=='none' or gett[s][gett[s].index('每箱规格')+1]=='':
			global findell
			findell=[]
			for y in range(len(zxtype)):
				e=[]
				if gett[s][gett[s].index('内盒')+1]==zxtype[y][0]:
					
					e.append(gett[s][gett[s].index('型号')+1])
					e.append(zxtype[y][1])
					e.append(zxtype[y][0])
					findell.append(e)
					e=[]
			if gett[s][gett[s].index('内盒')+1]=='':
				
				e.append(gett[s][gett[s].index('型号')+1])
				e.append('')
				e.append('')
				findell.append(e)
				e=[]
			print findell,'finfinfnifnifninfinfinifnfininfifnifn'
			# print need,'neeeeeeeeeeeeeeeeeeeeeeeeeeeeeeed'

#最终处理一匹配 
			if len(findell)==1:
				for z in range(len(zxtype)):
					if zxtype[z][0]==findell[0][2] and zxtype[z][1]==findell[0][1]:
						gett[s][gett[s].index('每箱规格')+1]=u'%s最终可能%s'%(zxtype[z][2],zxtype[z][1])
					try:
						gett[s][gett[s].index('毛重')+1]=zxtype[z][2]+gett[s][gett[s].indbex('净重')+1]
						gett[s][gett[s].index('总重量')+1]=gett[s][gett[s].index('毛重')+1]*gett[s][gett[s].index('数量')+1]
					except:
						print gett[s][gett[s].index('型号')+1],'zhongliang is woring'
					if findell[0][1]=='':
						gett[s][gett[s].index('每箱规格')+1]=''
						# gett[s][gett[s].index('每箱规格')+1]=zxtype[z][2]+'最终可能'+zxtype[z][1]

#最终处理二匹配
			elif len(findell)==2:
				a=[]
				b=[]
				wanquan=[]#全装箱
				alltype=[]
				print need
				print findell
				for k in range(len(findell)):
						alltype.append(findell[k][1])
						alltype.sort()
						alltype.reverse()
				print alltype,'typetype'
				zongshu=gett[s][gett[s].index('数量')+1]
				# print zongshu,'zonsgaid'
				# print zongshu//alltype[0]
				# try:					
				for x in range(int(zongshu//alltype[0])+1):
					for y in range(int(zongshu//alltype[1])+1):
							if alltype[0]*x+alltype[1]*y<=zongshu and zongshu-alltype[0]*x-alltype[1]*y<alltype[0]:
								global a
								b=[]
								b.append(x)
								b.append(y)
								b.append(zongshu-alltype[0]*x-alltype[1]*y)
								b.append(x+y)
							if b!=[] and b not in a:
					
							# print 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
								a.append(b)
								print alltype[0]*x+alltype[1]*y-zongshu
				print a,'aaaaaaaaaaaaaa'
				finaldeal(2,a,alltype)
				finalwrite(finalway,alltype,gett)
				# except:
				# 	print'lake'
				



#最终处理三匹配
			elif len(findell)==3:
				a=[]
				b=[]
				wanquan=[]#全装箱
				alltype=[]
				print need
				print findell
				for k in range(len(findell)):
						alltype.append(findell[k][1])
						alltype.sort()
						alltype.reverse()
				print alltype,'typetype'
				zongshu=gett[s][gett[s].index('数量')+1]
				print zongshu,'zonsgaid'
				# print zongshu//alltype[0]
				try:
					for x in range(int(zongshu//alltype[0])+1):
						for y in range(int(zongshu//alltype[1])+1):
							for z in range(int(zongshu//alltype[2])+1):
								if alltype[0]*x+alltype[1]*y+alltype[2]*z<=zongshu and zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z<alltype[0]:
									global a
									b=[]
									b.append(x)
									b.append(y)
									b.append(z)
									b.append(zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z)
									b.append(x+y+z)
								if b!=[] and b not in a:
						
								# print 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
									a.append(b)
									print alltype[0]*x+alltype[1]*y+alltype[2]*z-zongshu

					print a,'aaaaaaaaaaaaaa'
					finaldeal(3,a,alltype)
					finalwrite(finalway,alltype,gett)
				except:
					print 'lake'
				
#最终处理四匹配
			elif len(findell)==4:
				a=[]
				b=[]
				wanquan=[]#全装箱
				alltype=[]
				print need
				print findell
				for k in range(len(findell)):
						alltype.append(findell[k][1])
						alltype.sort()
						alltype.reverse()
				print alltype,'typetype'
				zongshu=gett[s][gett[s].index('数量')+1]
				try:
					print zongshu,'zonsgaid'
					print zongshu//alltype[0]
					for x in range(int(zongshu//alltype[0])+1):
						for y in range(int(zongshu//alltype[1])+1):
							for z in range(int(zongshu//alltype[2])+1):
								for j in range(int(zongshu//alltype[3])+1):
									if alltype[0]*x+alltype[1]*y+alltype[2]*z+alltype[3]*j<=zongshu and zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z-alltype[3]*j<alltype[0]:
										global a
										b=[]
										b.append(x)
										b.append(y)
										b.append(z)
										b.append(j)
										b.append(zongshu-alltype[0]*x-alltype[1]*y-alltype[2]*z-alltype[3]*j)
										b.append(x+y+z+j)
									if b!=[] and b not in a:
						
								# print 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
										a.append(b)
										print alltype[0]*x+alltype[1]*y+alltype[2]*z+alltype[3]*j-zongshu
					print a,'aaaaaaaaaaaaaa'
					finaldeal(4,a,alltype)
					finalwrite(finalway,alltype,gett)
				except:
					print 'lake'





	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')

	for n in range(len(gett)):#将相符的内容显示出来

		for i in range(len(gett[n])):#数据逐行写入excel
			# print len(gett[n])
			if is_chinese(gett[n][i]):
				gett[n][i].encode('utf-8')
			elif is_num(gett[n][i])==1:
				if math.modf(gett[n][i])[0]==0 or gett[n][i]==0:#获取数字的整数和小数
					gett[n][i]=int(gett[n][i])#将浮点数化成整数
			sheet1.write(n,i,gett[n][i])
		
	book.save('4.xls')#存储excel
	book = xlrd.open_workbook('4.xls')


def outtype(content):
	zxsmalltype=[]
	zxzongtype=[]
	flag=0
	global zxtype
	global zxzongtype1
	filename1(content)
	#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		nrows=table.nrows
		for n in range(1,nrows):
			#获取单行内容

			a=table.row_values(n)
			b=table.row_values(0)
			c=[]
			d=[]#新增的数组
			sheet = workbook.sheet_by_index(name)
			xfx = sheet.cell_xf_index(n, 0)
			xf = workbook.xf_list[xfx]
			bgx = xf.background.pattern_colour_index
			for l in range(len(a)): 
				c.append(b[l])
				c.append(a[l])
				c.append(bgx)
			
			# print c
			if type(c[c.index('内盒')+1])  is not types.FloatType :
				# print c[-1],'xxxxxxxxxxxxxxxxxxxxxxx!!!!!!!!!!!!!!!!!!!'
				c[c.index('内盒')+1]=c[c.index('内盒')+1].encode('utf-8').replace('X','*')
				c[c.index('内盒')+1]=unicode(c[c.index('内盒')+1],'utf-8')
			if c[-1]==11:
				zxsmalltype.append(c[c.index('内盒')+1])
				zxsmalltype.append(c[c.index('数量')+1])
				if 'CM' not in c[c.index('每箱规格')+1]:
					c[c.index('每箱规格')+1]=c[c.index('每箱规格')+1]+'CM'
				zxsmalltype.append(c[c.index('每箱规格')+1])
				zxsmalltype.append(c[c.index('重量')+1])
				zxzongtype.append(zxsmalltype)
				zxsmalltype=[]
	for i in zxzongtype:
		if i not in zxzongtype1  and i[0]!='' and i[1]!='':
			zxzongtype1.append(i)

	print zxzongtype1,'zxzongtype1'
	for x in range (0,len(zxzongtype1)-2):
		if x>=len(zxzongtype1)-2 :
			break
		zxtype.append(zxzongtype1[x])
		for y in range (x+1,len(zxzongtype1)-1):
			if y>=len(zxzongtype1)-1:
				break
			# print len(zxzongtype1),y,x
			if flag==1:#删除之后要倒退一位
				if  zxzongtype1[x][0]==zxzongtype1[y-1][0] and zxzongtype1[x][1]==zxzongtype1[y-1][1]:
					del zxzongtype1[y-1]
				else : flag=0
			if zxzongtype1[x][0]==zxzongtype1[y][0] and zxzongtype1[x][1]==zxzongtype1[y][1] and flag==0:	
				# print zxzongtype1[y]			
				del zxzongtype1[y]
				flag=1

	print zxzongtype1,'zxzongtype1'
	
	ling=[[102,15,u'30\xd720\xd710CM',1.3],[104,10,u'30\xd720\xd710CM',1.3],[105,10,u'30\xd720\xd710CM',1.3],[106,7,u'30\xd720\xd710CM',1.3],[107,6,u'30\xd720\xd710CM',1.3],[115,10,u'30\xd720\xd710CM',1.3]]
	print len(zxtype),'bbbbbbbbbbbbbbbbbbbbbbbbbb'
	for x in range(len(ling)):
		zxtype.append(ling[x])
		print len(zxtype)
	print zxtype,'b3333333333333333333333333333333333333333333333333333'
	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	for x in range(len(zxtype)):
		for y in range(0,4):
			if is_chinese(zxtype[x][y]):
				zxtype[x][y].encode('utf-8')
			# elif not four[i] nd four[i]!=0:
			# 	print "空值",
			elif is_num(zxtype[x][y])==1:
				if math.modf(zxtype[x][y])[0]==0 or zxtype[x][y]==0:#获取数字的整数和小数
					zxtype[x][y]=int(zxtype[x][y])#将浮点数化成整数
			sheet1.write(x,y,zxtype[x][y])		
	book.save('zxtype.xls')#存储excel
	book = xlrd.open_workbook('zxtype.xls')


def outt(content):
	lasttype=lastgett(gett)
	print lasttype,'gggtttttgtgtgtgtgtgtgtgttgtggtgtg'
	filename(content)
	three=[]
	four=[]
	# print gett
	xinghao=[]#所有存在three中的型号元素
	line=1
	flag=[]
	a=[]
	fla=0
	Sheetname=workbook.sheet_names()
	for name in range(len(Sheetname)):
		table = workbook.sheets()[name]
		nrows=table.nrows
	
		for n in range(nrows):
			#获取单行内容
			a=table.row_values(n)
			lena=len(a)
			for k in range(len(lasttype)):#将所有符合内容的部分筛选出来
				for x in range(len(a)):
					fla=0
					for m in range(len(lasttype[k])):						
						if a[x]==lasttype[k][m]:
							three.append(lasttype[k][m+1])
							if x not in flag:
								flag.append(x)#记录3中有数据的列
								fla=0
						elif a[x] not in lasttype[k] and fla==0:	
							three.append('')
							fla=1

	flag.sort()
	print flag
	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')


	# print a
	for i in range(len(a)):
		sheet1.write(0,i,a[i])#在第一行写入原先的数据

	print three
	print needitem,'neddde'

	for y in range(len(three)):
		for  x in range(len(needitem)):
			if needitem[x]==three[y]:
				xinghao.append(y)
	print xinghao,'xinxinxixnixxnxin'
	xinghao.sort()
	print xinghao[-1]
	print xinghao,'xinxinxixnixxnxin'

	# for n in range(0,len(three)-len(flag),len(flag)):
	# 	# print(three[n:len(flag)+n])
	# 	four=three[n:len(flag)+n]
	for x in range(len(xinghao)):
		if xinghao[x]==xinghao[-1]:#判断是否是最后一个区间
			print xinghao[-1]
			four=three[xinghao[x]:len(three)]
		else: four=three[xinghao[x]:xinghao[x+1]]
		print four
		
		guige=[]
			# print four
		for i in range(len(four)):
			if is_chinese(four[i]):
				four[i].encode('utf-8')
			# elif not four[i] and four[i]!=0: 
			# 	print "空值",
			elif is_num(four[i])==1:
				if math.modf(four[i])[0]==0 or four[i]==0:#获取数字的整数和小数
					four[i]=int(four[i])#将浮点数化成整数
			sheet1.write(line,i,four[i])
		line=line+1		
	book.save('jieguo.xls')#存储excel
	book = xlrd.open_workbook('jieguo.xls')
	check()

def readexcel3(content):
	filename(content)
	#获取所有的sheet
	for k in range(len(needitem)):#查找2里面所有和1相符的内容
		Sheetname=workbook.sheet_names()
		for name in range(len(Sheetname)):
			table = workbook.sheets()[name]
			nrows=table.nrows
			for n in range(nrows):
				#获取单行内容
				a=table.row_values(n)
				b=table.row_values(0)
				c=[]
				for l in range(len(a)):
					c.append(b[l])
					c.append(a[l])
					# print c,'ccccccccccccccccxssxsxsx'
			

				if needitem[k] in c: 
					keren.append(c)
					break
				else :
					if type(c[1]) is not types.FloatType :
						print'stringtype',c[1]
						try:
							c1=c[1].encode("utf-8")  
							needitemk=str(needitem[k])
							print needitemk,type(needitemk)
							print c1,type(c1)
							if needitemk in c1:
								try:
									c[c.index('型号')+1]=int(needitemk)
								except:
									print 'ctype is not int'
								keren.append(c)
								print 'sucusuihciauhciuhancao'
								break
						except:
							print 'is str'
							print c[1],needitem[k],'????????Ww'
							break


	print keren,'kerererenrnenrnenrnenrnernn'
	for x in range(len(gett)):
		for y in range(len(keren)):    
			if (gett[x][1]==keren[y][keren[y].index('型号')+1]):
				mid=keren[y][keren[y].index('型号')+1:len(keren[y])]
				gett[x]=gett[x]+mid
				gett[x].append(unicode('金额','utf-8'))
				gett[x].append(gett[x][gett[x].index('需求数量')+1]*mid[mid.index('单价')+1])						


	book = Workbook()
	sheet1 = book.add_sheet('Sheet 1')
	for n in range(len(gett)):#将相符的内容显示出来
		for i in range(len(gett[n])):#数据逐行写入excel
			if is_chinese(gett[n][i]):
				gett[n][i].encode('utf-8')
			# elif not gett[n][i] and gett[n][i]!=0:
			# 	print "空值",
			elif is_num(gett[n][i])==1:
				if math.modf(gett[n][i])[0]==0 or gett[n][i]==0:#获取数字的整数和小数
					gett[n][i]=int(gett[n][i])#将浮点数化成整数
			sheet1.write(n,i,gett[n][i])
		
	book.save('6.xls')#存储excel
	book = xlrd.open_workbook('6.xls')




def super():
	readexcel('型号和数量')
	outtype('zx')
	chuchong('客户型号和单价')
	readexcel2('zx')
	readexcel3('10')
	outt('目标装箱表格样式')



	


#主流程
def showmenu():
	prompt = """
(E)xcel文件打开
(W)打开excel并写入文件
(R)对比excel并导出
(f)获取客人编号和单价
(v)根据excel需求的内容导出
(c)计算
(s)super
(Q)uit

Enter choice: """
	done=False
	while not done:
		y=raw_input(prompt).strip()[0].lower()#获取字符串第一位并转小写
		if y=='q':done = True
		elif y=='w':write_excel()
		elif y=='e':readexcel()
		elif y=='r':readexcel2()
		elif y=='f':readexcel3()
		elif y=='v':outt()
		elif y=='c':cal()
		elif y=='s':super()
		else:
			print "请输入正确的选择"


if __name__ == "__main__":
	global needitem,neednum,need,gett,keren,alll,zxtype,zxzongtype1
	needitem=[]
	neednum=[]
	need=[] 
	gett=[]
	keren=[]
	alll=[]
	zxtype=[]
	zxzongtype1=[]
	super()
	# readpy()
