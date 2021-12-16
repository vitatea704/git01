import xlrd
import xlwt

# 需处理的文件名
file = '21设备发货明细表.xls'
# 读取文件
data = xlrd.open_workbook(file)
# 读取要处理的表名
#table = data.sheet_by_name("明细210914")
table = data.sheet_by_name("明细202110月")

nrows = table.nrows  # 行数
ncols = table.ncols  # 列数

print(nrows - 1)
#print(ncols)
# 添加表头
workbook = xlwt.Workbook(encoding='utf-8')
new_sheet = workbook.add_sheet('建设银行')

data = input('输入你想要筛选的数据,format(#建设银行)\n')
# data1 = input('输入第几列，format(3)\n')

rank_list = []
for i in range(1, nrows):
    if table.row_values(i)[5] == data:  # 筛选第几列就改 [1] 里的数字，数字从0开始起步
        rank_list.append(i)
print(rank_list)
# 写表头
for i in range(ncols):
    new_sheet.write(0, i, table.cell(0, i).value)

for i in range(len(rank_list)):
    for j in range(ncols):
        new_sheet.write(i + 1, j, table.cell(rank_list[i], j).value)

workbook.save('建设银行.xls')

##跟进第一天