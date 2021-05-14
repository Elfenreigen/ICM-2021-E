import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import random
import openpyxl
import pyecharts.options as opts
from pyecharts.charts import Bar3D

RID1=0.01
RID2=0.001
MAXNUM=365
weight=[0.3109,0.20132,0.19794,0.15936,0.13052]
cum_w=np.cumsum(weight)

data=pd.read_excel('table1.xlsx')
index1=list(data['FRI'])
index2=list(data['WRI'])
index3=list(data['FII'])
index4=list(data['ATI'])
index5=list(data['PRFT'])
index6=list(data['SSI'])
index7=list(data['GF'])
index8=list(data['HI'])
index=[index1,index2,index3,index4,index5,index6,index7,index8]
#index3 and index8 are negetaive indexes while the other are positive indexes

delta1_y=[index1[i+1]-index1[i] for i in range(9)]
delta2_y=[index2[i+1]-index2[i] for i in range(9)]
delta3_y=[index3[i+1]-index3[i] for i in range(9)]
delta4_y=[index4[i+1]-index4[i] for i in range(9)]
delta5_y=[index5[i+1]-index5[i] for i in range(9)]
delta6_y=[index6[i+1]-index6[i] for i in range(9)]
delta7_y=[index7[i+1]-index7[i] for i in range(9)]
delta8_y=[index8[i+1]-index8[i] for i in range(9)]
delta1_d=[i/365 for i in delta1_y]
delta2_d=[i/365 for i in delta2_y]
delta3_d=[i/365 for i in delta3_y]
delta4_d=[i/365 for i in delta4_y]
delta5_d=[i/365 for i in delta5_y]
delta6_d=[i/365 for i in delta6_y]
delta7_d=[i/365 for i in delta7_y]
delta8_d=[i/365 for i in delta8_y]
delta_d=[delta1_d,delta2_d,delta3_d,delta4_d,delta5_d,delta6_d,delta7_d,delta8_d]

def calculate():
	new_delta1=[]
	new_delta2=[]
	new_delta3=[]
	new_delta4=[]
	new_delta5=[]
	new_delta6=[]
	new_delta7=[]
	new_delta8=[]
	new_delta=[new_delta1,new_delta2,new_delta3,new_delta4,new_delta5
	,new_delta6,new_delta7,new_delta8]
	for i in range(9):
		record1=[]
		record2=[]
		record3=[]
		record4=[]
		record5=[]
		record6=[]
		record7=[]
		record8=[]
		record=[record1,record2,record3,record4,record5,record6,record7,record8]
		for j in range(MAXNUM):
			tmp=[0]*8
			r=random.random()
			priority_index=0
			if r>cum_w[0]:
				for k in range(4):
					if r<=cum_w[k+1]:
						priority_index=k+1
						break
						
			if priority_index==0: #sustainability
				tmp[0]=abs(delta_d[0][i])*RID1+abs(index[0][i])*RID2
				tmp[1]=abs(delta_d[1][i])*RID1+abs(index[1][i])*RID2
				for m in range(8):
					if m!=0 and m!=1 and m!=2 and m!=7:
						tmp[m]=-abs(delta_d[m][i])*RID1/4
						-abs(index[m][i])*RID2/4
					if m==2 or m==7:
						tmp[m]=abs(delta_d[m][i])*RID1/4
						+abs(index[m][i])*RID2/4
			else:
				if priority_index==1: #enquity
					tmp[2]=-abs(delta_d[2][i])*RID1-abs(index[2][i])*RID2
					tmp[5]=abs(delta_d[5][i])*RID1+abs(index[5][i])*RID2
					tmp[7]=-abs(delta_d[7][i])*RID1-abs(index[7][i])*RID2
					for m in range(8):
						if m!=2 and m!=5 and m!=7:
							tmp[m]=-abs(delta_d[m][i])*RID1/4
							-abs(index[m][i])*RID2/4
				else:
					if priority_index==2: #efficiency
						tmp[3]=abs(delta_d[3][i])*RID1+abs(index[3][i])*RID2
						for m in range(8):
							if m!=3 and m!=2 and m!=7:
								tmp[m]=-abs(delta_d[m][i])*RID1/4
								-abs(index[m][i])*RID2/4
							if m==2 or m==7:
								tmp[m]=abs(delta_d[m][i])*RID1/4
								+abs(index[m][i])*RID2/4
					else:
						if priority_index==3: #profitability
							tmp[4]=abs(delta_d[4][i])*RID1+abs(index[4][i])*RID2
							for m in range(8):
								if m!=4 and m!=2 and m!=7:
									tmp[m]=-abs(delta_d[m][i])*RID1/4
									-abs(index[m][i])*RID2/4
								if m==2 or m==7:
									tmp[m]=abs(delta_d[m][i])*RID1/4
									+abs(index[m][i])*RID2/4
						else: #globalization
							tmp[6]=abs(delta_d[6][i])*RID1+abs(index[6][i])*RID2
							for m in range(8):
								if m!=6 and m!=2 and m!=7:
									tmp[m]=-abs(delta_d[m][i])*RID1/4
									-abs(index[m][i])*RID2/4
								if m==2 or m==7:
									tmp[m]=abs(delta_d[m][i])*RID1/4
									+abs(index[m][i])*RID2/4
									
			for m in range(8):
				record[m].append(tmp[m])
				
		for m in range(8):
			new_delta[m].append(sum(record[m]))
	return new_delta

############################
#RID1=0.01,RID2=0.001,Roulette Wheel Selection
new_delta=calculate()
for m in range(8):		
		print(new_delta[m])

print(sum(new_delta[7]))
wb=openpyxl.Workbook()
ws=wb.active
for r in range(len(new_delta)):
	for c in range(len(new_delta[0])):
		ws.cell(c+1,r+1).value=new_delta[r][c]
wb.save('Japan.xlsx')

#########################
#sensitivity analysis
sensibility=[]
for i in range(20):
	RID1=0.01+i*0.005
	for j in range(20):
		RID2=0.001+0.0005*j
		new_delta=calculate()
		value=sum(new_delta[1])
		new_arr=[RID1,RID2,value]
		sensibility.append(new_arr)
		
file_=open('file.txt','w')
file_.write(str(sensibility))
file_.close()

test=[[d[1], d[0], d[2]] for d in sensibility]


c=(Bar3D(init_opts=opts.InitOpts(width="1600px", height="800px"))
.add(
	series_name="",
	data=test,
	xaxis3d_opts=opts.Axis3DOpts(name='RID2'),
	yaxis3d_opts=opts.Axis3DOpts(name='RID1'),
	zaxis3d_opts=opts.Axis3DOpts(name='WRI')
)
.render('plot3D.html'))



