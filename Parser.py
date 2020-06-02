import pandas as pd
import numpy as np
from operator import itemgetter
import easygui

path = easygui.fileopenbox()
#read excel to {'Mapped AD IMS for testing-new': dataframe, 'Mapped AD IMS for testing-old': dataframe}
data = pd.read_excel(path, sheet_name = ['Mapped AD IMS for testing-new', 'Mapped AD IMS for testing-old'], usecols = ['SCOPE', 'Milestone', 'Application', 'ACIO', "Status_Date"])
date = (data['Mapped AD IMS for testing-new'].at[1,'Status_Date']).strftime("%d_%b_%Y")
#remove unnecessary data from the dataframes
#Mapped AD IMS for testing-new
rowIndex = []
for index, row in data['Mapped AD IMS for testing-new'].iterrows():
	if row['Milestone'] != 'INITIATIVE' and row['Milestone'] != 'APPLICATION':
		if pd.isnull(data['Mapped AD IMS for testing-new'].at[index,'Milestone']) == False or (row['SCOPE'] != 'ARCHIVE' and row['SCOPE'] != 'DESCOPE'):
			rowIndex.append(index)
data['Mapped AD IMS for testing-new'] = data['Mapped AD IMS for testing-new'].drop(rowIndex)
#Mapped AD IMS for testing-old
rowIndex = []
for index, row in data['Mapped AD IMS for testing-old'].iterrows():
	if row['Milestone'] != 'INITIATIVE' and row['Milestone'] != 'APPLICATION':
		if pd.isnull(data['Mapped AD IMS for testing-old'].at[index,'Milestone']) == False or (row['SCOPE'] != 'ARCHIVE' and row['SCOPE'] != 'DESCOPE'):
			rowIndex.append(index)
data['Mapped AD IMS for testing-old'] = data['Mapped AD IMS for testing-old'].drop(rowIndex)

#Build dataframes
#Create 3 lists, new, old, and both
new = []
old = []
both = []
acio = {}
#Compare new list
values = set(data['Mapped AD IMS for testing-old']['Application'])
data['Mapped AD IMS for testing-new']['Match'] = data['Mapped AD IMS for testing-new']['Application'].isin(values)
#Compare old list
values = set(data['Mapped AD IMS for testing-new']['Application'])
data['Mapped AD IMS for testing-old']['Match'] = data['Mapped AD IMS for testing-old']['Application'].isin(values)
#Add data to correct lists
for index, row in data['Mapped AD IMS for testing-new'].iterrows():
	if row['Match'] == False:
		if row['Milestone'] == 'APPLICATION':
			new.append((row['Application'],row['ACIO'], 'APP'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][0] = acio[row['ACIO']][0] + 1
			else:
				acio[row['ACIO']] = [1,0]
		else:
			new.append((row['Application'],row['ACIO'], 'INIT'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][1] = acio[row['ACIO']][1] + 1
			else:
				acio[row['ACIO']] = [0,1]
		
	if row['Match'] == True:
		if row['Milestone'] == 'APPLICATION':
			both.append((row['Application'],row['ACIO'], 'APP'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][0] = acio[row['ACIO']][0] + 1
			else:
				acio[row['ACIO']] = [1,0]
		else:
			both.append((row['Application'],row['ACIO'], 'INIT'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][1] = acio[row['ACIO']][1] + 1
			else:
				acio[row['ACIO']] = [0,1]


for index, row in data['Mapped AD IMS for testing-old'].iterrows():
	if row['Match'] == False:
		if row['Milestone'] == 'APPLICATION':
			old.append((row['Application'],row['ACIO'], 'APP'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][0] = acio[row['ACIO']][0] + 1
			else:
				acio[row['ACIO']] = [1,0]
		else:
			old.append((row['Application'],row['ACIO'], 'INIT'))
			if row['ACIO'] in acio:
				acio[row['ACIO']][1] = acio[row['ACIO']][1] + 1
			else:
				acio[row['ACIO']] = [0,1]

#convert to dataframes
new = pd.DataFrame(sorted(sorted(new,key=itemgetter(1)), key=itemgetter(2)), columns = ['Application/Initiative', 'ACIO', 'Type'])
old = pd.DataFrame(sorted(sorted(old,key=itemgetter(1)), key=itemgetter(2)), columns = ['Application/Initiative', 'ACIO', 'Type'])
both = pd.DataFrame(sorted(sorted(both,key=itemgetter(1)), key=itemgetter(2)), columns = ['Application/Initiative', 'ACIO', 'Type'])
acio = pd.DataFrame(data = acio, index = ['Application', 'Initiative']).T

#write to excel
with pd.ExcelWriter('MappingComparison_'+ date + '.xlsx') as writer:  
	new.to_excel(writer, sheet_name = 'New',index = False)
	both.to_excel(writer, sheet_name = 'Both',index = False)
	old.to_excel(writer, sheet_name = 'Old',index = False)
	acio.to_excel(writer, sheet_name = 'Acio')

print('Finished')