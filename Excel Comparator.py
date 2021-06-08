import pandas as pd

newFile = '10_12 - Copy.xlsx'
oldFile = '10_12.xlsx'

file = pd.ExcelFile(newFile)
sheet = file.sheet_names[0]                                                # load sheet(0 indexed) you want to compare

df_new = pd.read_excel(newFile, sheet_name=sheet)                          # load new file data
df_old = pd.read_excel(oldFile, sheet_name=sheet)                          # load old file data

sharedCols = list(set(df_old.columns).intersection(df_new.columns))        # get shared columns in both DFs
dup_new = df_new.iloc[:, 1]                                                # get all the values in the column indexed 1(using which you want to identify data)
dup_old = df_old.iloc[:, 1]

'''
below generator function used in case values are not unique in
column used to compare both data. If XY occurs 3 times then this function generates XY, XY_1, XY_2.
'''
def rename_duplicates(dup_servers):                                         
    seen = {}                                                              
    for x in dup_servers:
        if x in seen:
            seen[x] += 1
            yield "%s_%d" % (x, seen[x])
        else:
            seen[x] = 0
            yield x

df_new['dup'] = list(rename_duplicates(dup_new))                           # list generated as output of function is added as   
df_old['dup'] = list(rename_duplicates(dup_old))                           # a new column to both the DFs

df_join = df_old.merge(df_new, on='dup', how='outer',suffixes=('_old','_new'), indicator=True)   # perform an outer join

deleted_value = df_join.loc[df_join['_merge'] == 'left_only', 'dup']       # this list is generated as its needed in the end so, deleted
deleted_data = df_old[df_old['dup'].isin(deleted_value)]                   # data is fetched this way using above list

new_value = df_join.loc[df_join['_merge'] == 'right_only', 'dup']          # this list is needed in the end
new_data = df_new[df_new['dup'].isin(new_value)]

common = df_join.loc[df_join['_merge'] == 'both', 'dup']                   # list of 'dup' column values that is common in both DFs

df_new.set_index('dup', inplace=True)
df_old.set_index('dup', inplace=True)
new_data.set_index('dup', inplace=True)
deleted_data.set_index('dup', inplace=True)

dfDiff = pd.DataFrame(columns=sharedCols, index=common)
for row in dfDiff.index:                                                   # loop through each row in dfDiff
    for col in sharedCols:                                                 # loop through each column in dfDiff
        value_old = df_old.loc[row, col]                                   # get cell value from old df
        value_new = df_new.loc[row, col]                                   # get cell value from new df
        if value_old == value_new:
            dfDiff.loc[row, col] = value_new                               # if both values are same that its added as it is in dfDiff
        else:
            dfDiff.loc[row, col] = ('{}→{}').format(value_old, value_new)  # otherwise it gets added in a format "old->new"

only_changed = pd.DataFrame()                                              # an empty df that will contain only those rows which have a change
for col in sharedCols:                                                     # rows are selected using each column one by one
    only_changed = only_changed.append(dfDiff[dfDiff[col].apply(str).str.contains('→')])     # Only select rows cotaining ->, rows can be duplicate as this operation is done for each column of a row

changed_data = only_changed.drop_duplicates()                              # to remove duplicate rows 
dfList = [new_data, changed_data, deleted_data]                            
finalDF = pd.concat(dfList)                                                # combine new, changed and deleted data

# Write to excel
fname = '{} Compared.xlsx'.format(sheet)
writer = pd.ExcelWriter(fname, engine='xlsxwriter')
finalDF.to_excel(writer, sheet_name=sheet, index=False)
workbook = writer.book
worksheet = writer.sheets[sheet]

# Apply formatting
grey_fmt = workbook.add_format({'font_color': '#999999'})                  # format for deleted data
highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color': '#B1B3B3'})   # format for changed data
new_fmt = workbook.add_format({'font_color': '#32CD32', 'bold': True})     # format for new data
worksheet.conditional_format('A1:ZZ1000', {'type': 'text', 'criteria': 'containing', 'value': '→', 'format': highlight_fmt})

for rownum in range(finalDF.shape[0]):
    row = finalDF.index[rownum]
    for x in new_value:                                                    # loop through new ''dup' values
        if x == row:
            worksheet.set_row(rownum+1, 15, new_fmt)                       # apply format for new data if they match
    for y in deleted_value:                                                # loop through deleted 'dup' values
        if y == row:
            worksheet.set_row(rownum+1, 15, grey_fmt)                      # apply format for deleted data if they match

writer.save()
print('\nDone.\n')