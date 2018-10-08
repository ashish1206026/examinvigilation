#!python3
#optimization_invigilation.py - a test program for creating automatic excel files for the invigilators

import os,openpyxl,xlsxwriter
os.chdir('D:\\python\\Sonargaon_University')
wb = openpyxl.load_workbook('final invigilation.xlsx')
sheet = wb['Sheet1']
columns = list(sheet.iter_cols())
rows = list(sheet.iter_rows())
#rows = rows[2:]
col_date=columns[1]

#Checing for some formatting function
#workbook=xlsxwriter.Workbook('A.xlsx')
#worksheet=workbook.add_worksheet()
#r=1
#c=1
#for i in col_date:
	#print(i.value)
	#format_date=workbook.add_format({'num_format': 'd-mmm-yyyy'})
	#worksheet.write(r,c,i.value,format_date)
	#c=c+1
#workbook.close()


intro = rows[0] 
for i in range(0,len(columns)):
	columns[i] = list(columns[i][2:])
invig1 = columns[12][1:]
invig2 = columns[13][1:]
invig3 = columns[14][1:]
invig = {}
invignames = []

# The following block of code will be used for formatting the rows and columns for later sorting and also 
# will automatically create a dictionary with the invig. names as keys and list of sl no as values
for i in range(0,len(invig1)):
	if(invig1[i].value != None):
		invig1[i].value = invig1[i].value.title()
		j = invig1[i].value
		if((j in invig.keys()) == False):
			invig[j] = []
			invig[j].append(i+1)
			invignames.append(j)
		else:
			invig[j].append(i+1)
	if(invig2[i].value != None):
		invig2[i].value = invig2[i].value.title()
		j = invig2[i].value
		if((j in invig.keys()) == False):
			invig[j] = []
			invig[j].append(i+1)
			invignames.append(j)
		else:
			invig[j].append(i+1)	
	if(invig3[i].value != None):
		invig3[i].value = invig3[i].value.title()
		j = invig3[i].value
		if((j in invig.keys()) == False):
			invig[j] = []
			invig[j].append(i+1)
			invignames.append(j)
		else:
			invig[j].append(i+1)	

# The following code will generate individual excel files for all the invigilators automatically
for name in invignames:
	workbook=xlsxwriter.Workbook(name+'.xlsx')
	worksheet=workbook.add_worksheet()
	format_date=workbook.add_format({'num_format':'dd-mm-yy'})
	r=1
	c=0
	for i in intro:
		worksheet.write(r,c,i.value)
		c=c+1
	for i in invig[name]:
		c=0
		r=r+1
		for j in rows[i]:
			if(c==1):
				d=str(j.value)
				e=''
				for k in range(0,10):
					e=e+d[k]
				print(e)
				worksheet.write(r,c,e)
			else:
				worksheet.write(r,c,j.value)
			c=c+1
	workbook.close()
