import pandas as pd

newFile = 'Production Servers.xlsx'
oldFile = 'Production Servers old.xlsx'

df_new = pd.read_excel(newFile, sheet_name=5)                           # load sheet(0 indexed) you want to compare
df_old = pd.read_excel(oldFile, sheet_name=5)


df_join = df_old.merge(df_new,                                          # left is old df and right is new df
                       how='outer',                                     # perform an outer join so that no entries are excluded
                       on=['Server','Application'],                     # on columns necessary(in this case they form a unique key)
                       suffixes=('_old','_new'),                       
                       indicator=True)                                  # this attribute creates a column '_merge' telling a row(key) exists in which df

new_data = df_join[df_join['_merge'] == 'right_only']                   # df of only new data
deleted_data = df_join[df_join['_merge'] == 'left_only']                # df of only deleted data

combinedDF = pd.concat([new_data, deleted_data])
finalDF = combinedDF.iloc[:, [0,1]]                                     # this selects all rows and excludes the indicator column(index=2 here)

fname = 'Software Compared.xlsx'
writer = pd.ExcelWriter(fname, engine='xlsxwriter')
finalDF.to_excel(writer, sheet_name='Software', index=False)
workbook = writer.book
worksheet = writer.sheets['Software']

new_fmt = workbook.add_format({'font_color': '#32CD32', 'bold': True})  # formatting for new data
grey_fmt = workbook.add_format({'font_color': '#999999'})               # formatting for deleted data


for rownum in range(finalDF.shape[0]):
    if combinedDF.iloc[rownum, 2] == 'right_only':                      # this checks indicator column value
        worksheet.set_row(rownum+1, 15, new_fmt)                        # this formatting is applied if its in new(right) df
    else:
        worksheet.set_row(rownum+1, 15, grey_fmt)                       # this fomratting is applied if its in old(left) df

writer.save()
print('\nDone.\n')