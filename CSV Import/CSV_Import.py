import os
import pandas as pd

file = 'spreadsheet.xlsx'

xl = pd.ExcelFile(file)

print(xl.sheet_names)

df1 = xl.parse(0)

textFile = open('sql.txt', 'w+')

for i in df1.index:
	textFile.write('if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3712) begin update mi.RangeDetail set Name = \'{1} - {0}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3712 end \n'.format(df1['HRRef'][i], df1['FullName'][i]))
textFile.write('\n')
textFile.write('<----------------------							----------------------> \n')
textFile.write('\n')
df2 = xl.parse(1)

for i in df2.index:
	textFile.write('if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3715) begin update mi.RangeDetail set Name = \'{1} - {0}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3715 end \n'.format(df2['HRRef'][i], df2['FullName'][i]))
textFile.write('\n')
textFile.write('<----------------------							----------------------> \n')
textFile.write('\n')
df3 = xl.parse(2)

for i in df3.index:
	textFile.write('if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3714) begin update mi.RangeDetail set Name = \'{1} - {0}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3714 end \n'.format(df3['HRRef'][i], df3['FullName'][i]))
textFile.write('\n')
textFile.write('<----------------------							----------------------> \n')
textFile.write('\n')
df4 = xl.parse(3)

for i in df4.index:
	textFile.write('if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3713) begin update mi.RangeDetail set Name = \'{1} - {0}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3713 end \n'.format(df4['HRRef'][i], df4['FullName'][i]))

#df5 = xl.parse(4)

#for i in df5.index:
#	textFile.write('if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3716) begin update mi.RangeDetail set Name = \'{0} - {1}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3716 end \n'.format(df5['HRRef'][i], df5['FullName'][i]))
textFile.write('\n')
textFile.write('<----------------------							----------------------> \n')
textFile.write('\n')
df6 = xl.parse(5)

for i in df6.index:
	ActiveJobQuery = 'if exists (select 1 from mi.RangeDetail where Name = \'{1}\' and RangeID = 3710) begin update mi.RangeDetail set Name = \'{0} - {1}\', UpdateDate = GetDate(), UpdatePersonID = 1 where Name = \'{1}\' and RangeID = 3710 end \n'.format(df6['Job'][i], df6['Name'][i])
	ActiveJobQuery = ActiveJobQuery.replace('- -', ' -')
	ActiveJobQuery = ActiveJobQuery.replace('   ', ' ')
	#ActiveJobQuery = ActiveJobQuery.replace('  ', ' ')
	#ActiveJobQuery = ActiveJobQuery.replace('\' ', '\'')
	textFile.write(ActiveJobQuery)

textFile.close()

file = open('sql.txt', 'r')

fileData = file.read()


textFile = open('sql.txt', 'w')

textFile.write(fileData)