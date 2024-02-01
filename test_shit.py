from xlwt import Workbook

zoinks = ['a', 'aa', 'aaa', 'aaaa', 'aaaaa']
soinks = ['b', 'bb', 'bbb', 'bbbb', 'bbbbb']
numbers = ['1', '2', '3', '4', '5']
pain = numbers, zoinks, soinks

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

# Iterate over each list in the tuple
for i, lst in enumerate(pain):
    # Write each element in the list to a column in the current row
    for j, item in enumerate(lst):
        sheet1.write(i, j, item)

# Save the workbook
wb.save('xlwt example.xls')
