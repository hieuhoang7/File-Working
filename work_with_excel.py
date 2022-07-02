import openpyxl

f = openpyxl.load_workbook('excel_file.xlsx')
my_sheet1 = f['sheet1']
#print(my_sheet1['B2'].value) #xem du lieu cua o B2
for i in range(1,6): #lap toi dong du lieu thu n
    print(my_sheet1['B'+str(i)].value)

my_sheet1['B4'].value = 'Blackjack' #ghi du lieu vao file excel
f.save('excel_file.xlsx') #luu file
f.close()
