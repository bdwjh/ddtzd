# -*- encoding:utf-8 -*-

import re
import time
import sys
from glob import glob
import win32com.client
import os

class EasyExcel(object):
    '''
    需要安装模块：pip install pywin32
    需要导入模块：win32com.client
    '''
    def __init__(self, filename:str = '', show:bool = False):
        '''初始化 Excel 应用程序'''
        self.__file = filename                                            # 文件全路径
        self.__exists = False                                             # 文件是否存在
        self.__fileInfo = {}                                              # 文件扩展名
        self.__excel = win32com.client.Dispatch('Excel.Application')
        self.__excel.DisplayAlerts = False                                # 不显示提示信息
        self.__excel.Visible = show                                       # 显示窗体
        self.open()


    def open(self, file:str = ''):
        '''打开 Excel 文件'''
        if getattr(self, 'book', False):
            self.__book.Close()
        if file:
            self.__file = file
        self.__exists = os.path.isfile(self.__file)
        # 文件存在则打开，否则添加
        if not self.__file or not self.__exists:
            self.__book = self.__excel.Workbooks.Add()
            self.save()
        else:
            self.__book = self.__excel.Workbooks.Open(self.__file)
            self.__fileInfo = os.path.splitext(self.__file)
        self.get_sheet(1)                                        # 选择第一个工作表

    def add_sheet(self, sheetName:str = ''):
        '''添加一个新的工作表'''
        sheet = self.__book.Worksheets.Add()
        if sheetName:
            sheet.Name = sheetName
        #return sheet1
        self.__sheet = sheet

    def get_sheet(self, sheet):
        # assert sheet > 0, '工作表索引必须大于 0'
        # return self.__book.Worksheets(sheet)
        self.__sheet = self.__book.Worksheets(sheet)

    # def get_sheet_by_name(self, sheetName:str):
    #     for i in range(1, self.get_sheet_count()+1):
    #         sheet = self.__book.Worksheets(i)
    #         if sheetName == sheet.Name:
    #             #return sheet
    #             self.__sheet = sheet
    #             break
    #     #return None

    def get_sheet_count(self):
        '''获取工作表数量'''
        return self.__book.Worksheets.Count

    def reset(self):
        self.__excel    = None
        self.__book     = None
        self.__sheet    = None
        self.__file     = None
        self.__fileInfo = None

    def show(self, show = True):
        self.__excel.Visible = show

    def save(self, file:str = ''):
        '''保存 Excel 文件'''
        assert type(file) is str, "保存 Excel 文件名必须为字符型"
        
        # 同名或未设置文件名，则保存
        if not file or (self.__exists and file == self.__file):
            self.__book.Save()
            return

        # 当前文件与保存的文件扩展名是否相同，不同则改为相同
        fileInfo = os.path.splitext(file)
        if self.__fileInfo and (fileInfo[-1].upper() != self.__fileInfo[-1].upper()):
            file = fileInfo[0] + self.__fileInfo[-1]

        # 文件另存
        filePath = os.path.dirname(file)
        if not os.path.isdir(filePath):
            os.makedirs(filePath)

        self.__book.SaveAs(file)
        self.__file = file
        self.__fileInfo = os.path.splitext(self.__file)

    def quit(self):
        '''关闭 Excel 文件'''
        self.__book.Close(SaveChanges = 1)
        self.__excel.Quit()
        self.reset()

    def book_close(self):
        self.__book.Close(SaveChanges = 1)

    def get_cell(self, row=1, col=1):
        '''获取单元格对象'''
        assert row>0 and col>0, '要获取单元格所在行与列均须大于 0'
        return self.__sheet.Cells(row, col)

    def cell_value(self, row, col, value = None):
        '''获取或设置单元格值'''
        assert row>0 and col>0, '要获取单元格所在行与列均须大于 0'
        if value:
            self.get_cell(row, col).Value = value
        else:
            cell = self.get_cell(row, col)
            return self.get_cell(row, col).Value

    def get_rows_count(self):
        '''返回当前工作表最大行数'''
        # return self.__sheet.Rows.Count
        return self.__sheet.UsedRange.Rows.Count

    def get_cols_count(self):
        '''返回当前工作表最大列数'''
        # return self.__sheet.Cols.Count
        return self.__sheet.UsedRange.Columns.Count

    def get_row(self, row = 1):
        '''获取行对象'''
        assert row > 0, '获取行对象时行值必须大于 0'
        return self.__sheet.Rows(row)

    def select_row(self, row):
        '''选择行'''
        self.__sheet.Rows(row).Select

    def get_col(self, col = 1):
        '''获取列对象'''
        assert col > 0, '获取列对象时列值必须大于 0'
        return self.__sheet.Columns(col)

    def row_value(self, row = 1, value = None):
        '''获取或设置行值'''
        assert row > 0, '获取或设置行时行值必须大于 0'
        if value:
            self.get_row(row).Value = value
        else:
            return self.get_row(row).Value

    def get_range(self, row1, col1, row2, col2):
        return self.__sheet.Range(self.get_cell(row1, col1), self.get_cell(row2, col2))

    def find(self, range, search, option = True):
        '''在区域“range”中查找所有内容为“search”的记录，option为真时为完全相同，否则为包含'''
        find_value = []
        cell = range.Find(search, LookAt=option)
        if cell:                        # 成功查询则保存地址，用于结束查询测试
            addr = cell.Address
        while True:
            try:                        # 将find结果保存起来，同时包含findnext结果
                find_value.append(cell.Address)
            except AttributeError:      # 当Address属性不存在时即find没有结果
                break
            else:                       # 如果没有出错，意味查询结果保存完毕，进入下一次查询
                cell = range.FindNext(cell)
                if cell.Address == addr:    # 如果find已经找到，则这里总是能够找到
                    break                   # 当查找结果地址与第一次相同时结束查询
        return find_value


class TongZhiDan():
    __total      = 0        # 学生总数
    __chuli      = 0        # 处理学生记录数, 这是各报名信息记录数之和
    __zhizhi     = []
    __test       = []
    __re_study   = re.compile('学号：(\d{13})')     # 用于查找学号
    __re_subject = re.compile('\d{4}')             # 用于查找科目代码(试卷号)
    __require    = [
        '考场号', '考试日期', '考试时间', '科目', '学号', '座位号', '姓名'
    ]

    def __init__(self, zhizhi:str, test_info:str):
        self.__zhizhi.append(EasyExcel(zhizhi, False))
        self.__zhizhi.append(self.__zhizhi[0].get_rows_count())
        self.__zhizhi.append(self.__zhizhi[0].get_cols_count())
        # print(self.__zhizhi)
        for test in test_info:
            item    = []
            item.append(EasyExcel(test, False))
            item.append(item[0].get_rows_count())
            item.append(item[0].get_cols_count())
            item.append(self.get_fields_info(item[0], item[1], item[2], test))
            self.__test.append(item)

    def set_color(self, color):
        self.__color = color

    def end(self, save_path):
        # zhizhi, maxRows, maxCols = self.__zhizhi
        # zhizhi.save(r'D:\PythonVirtualEnv\DianDa\excel\考试通知单.xls')
        for test in self.__test:
            test_open, maxRows, maxCols, fields = test
            test_open.book_close()
        zhizhi, maxRows, maxCols = self.__zhizhi
        zhizhi.save(save_path)
        zhizhi.quit()
        # zhizhi.show(True)

        print('\n处理结束，文件保存到：{}'.format(save_path))
        print('共计找到考生数：{}'.format(self.__total))
        print('处理考试记录数：{}'.format(self.__chuli))
        #print('=============================================================')

    def start(self):
        zhizhi, maxRows, maxCols = self.__zhizhi
        # 遍历纸质通知单
        row = 2
        while row < maxRows:
            # 查找考生
            # row += 2
            cell_value = zhizhi.cell_value(row, 1)
            if cell_value is None:                # 跳过空白行
                continue
            search = self.__re_study.search(cell_value)
            if search:                            # 匹配成功，则查找相关科目考试信息
                student_num = search.group(1)
                # print('处理考生：', student_num)
                print('.', end='', flush=True)
                self.__total += 1
                # 查找并汇总考生报考编排明细
                test_info = self.get_test_info(student_num)
                # for a in test_info:    # 显示找到的第一条记录
                #     print(test_info[a])
                # exit()
                row += 2            # 跳过表头 试卷号、考试科目等
                while True:         # 遍历当前学生的所有考试科目
                    cell_value = zhizhi.cell_value(row, 1)
                    #if cell_value is None:                # 跳过空白行
                    #    continue
                    match = self.__re_subject.match(cell_value)
                    if match:       # 获取到科目代码（试卷号）
                        sub_code = match.group()
                        if sub_code in test_info:
                            self.__chuli += 1
                            # print(test_info)
                            # print(sub_code, self.__require[0])
                            zhizhi.cell_value(row, 4, test_info[sub_code][self.__require[0]])     # 考场号
                            zhizhi.cell_value(row, 5, test_info[sub_code][self.__require[5]])     # 座位号
                            zhizhi.cell_value(row, 6, test_info[sub_code][self.__require[1]])     # 考试日期
                            zhizhi.cell_value(row, 7, test_info[sub_code][self.__require[2]])     # 考试时间
                            zhizhi.get_cell(row, 7).NumberFormatLocal = "[$-x-systime]hh:mm:ss"   # 设置时间格式
                            if self.__color:
                                zhizhi.cell_value(row, 9, test_info[sub_code][self.__require[6]])     # 显示姓名
                                # 来点颜色
                                range = zhizhi.get_range(row, 4, row, 9)
                                range.Interior.ColorIndex = 6
                                range.Font.Color = 5
                    else:           # 不在匹配则处理结束，此时为考点名称处
                        break
                    row += 1        # 处理下一科目数据
                row += 3            # 处理下一个学生数据，此时为考点名称处，下移3行为学号位置



    def get_fields_info(self, test_open, maxRows, maxCols, test_file):
        '''获取字段名列表'''
        fields = {}
        for i in range(1, maxCols + 1):
            value = test_open.cell_value(1, i)
            if value is None:
                continue
            fields[value] = i
        # 检测必备字段是否存在，不存在则提示并退出程序
        for field in self.__require:
            if field not in fields:
                print('注意 “{}” 文件中不存在字段 “{}”，请修正'.format(test_file, field))
                wait = input('按回车键去修正...')
                exit()
        return fields

    def get_test_info(self, student_num):
        '''获取汇总的考生报考编排明细表详情'''
        info = {}
        # 考生报考编排明细表
        for test in self.__test:
            test_open, maxRows, maxCols, fields = test
            # print(test_open, maxRows, maxCols, fields)
            # 选择学号所在列
            study_field = fields[self.__require[4]]
            range = test_open.get_range(1, study_field, maxRows, study_field)
            item_addr = test_open.find(range, student_num)
            # details = self.get_test_details(test_open, item)
            # 遍历找到的记录地址，按行获取详细信息
            for addr in item_addr:
                details = {}
                row = int(addr.split('$')[-1])
                for field in self.__require:        # 只获取必须字段的内容
                    # print(fields, '===', field, '===', fields[field])
                    details[field] = test_open.cell_value(row, fields[field])
                info[details[self.__require[3]]] = details
        return info


                

def main():
    '''广播电视大学纸质考试通知单与报考编排明细表合并工具 V1.01
=============================================================================
使用方法一：
　　将包含考试通知单和省开、统设、计算机考生报考编排明细表的文件夹拖动到本程序上方，然后松手回答问题即可。
使用方法二：
　　直接运行程序，将要求输入通知单所在文件夹位置
注意事项：
　　1. 考生报考编排明细表表头必须包含字段：考场号、考试日期、考试时间、科目(仅代码)、学号、座位号
　　2. 考生报考编排明细表表头字段名空白将被跳过

    程序设计：王佳辉  bdwjh@163.com 2019-12-15  
    '''
    # print(main.__doc__)

    # zhizhi_excel = r"D:\PythonVirtualEnv\DianDa\excel\纸考考试通知单.xls"
    # test_excel   = [
    #     r"D:\PythonVirtualEnv\DianDa\excel\省开.xls",
    #     #r"D:\PythonVirtualEnv\DianDa\excel\省开考生报考编排明细表.xls",
    #     # r"D:\PythonVirtualEnv\DianDa\excel\统设考生报考编排明细表.xls",
    #     #r"D:\PythonVirtualEnv\DianDa\excel\计算机考生报考编排明细表.xlsx"
    #     ]
    tzd = TongZhiDan(zhizhi_excel, test_excel)
    co  = input('需要给你点颜色看看吗？(Y/N)：')
    if co and co.upper() != 'Y':
        co = False
        print('----------------------------没有颜色')
    else:
        co = True
        print('----------------------------有点颜色')
    tzd.set_color(co)                           # 来点颜色
    # tzd.get_test_info()
    print('=============================================================================')
    print('现在或品茶、或喝咖啡，享受一下快乐的生活吧！')
    print('=============================================================================')
    # time.sleep(5)
    print('正在处理',end='')
    tzd.start()
    tzd.end('{}\考试通知单-处理后.xls'.format(dir_path))

def get_dir_path():
    while True:
        dir_path = input('请输入考试通知单及考生报考编排明细表所在位置：')
        if dir_path:
            break
    return dir_path

def select_zhizhi():
    global zhizhi_excel
    zhi_pos = disp_dir()
    pos = input('请输入纸质考试通知单所在文件序号（回车默认 {}）：'.format(zhi_pos + 1))
    if pos:
        zhi_pos = int(pos)
    zhizhi_excel = test_excel.pop(zhi_pos)
    while True:
        disp_dir()
        pos = input('请输入要移除文件序号，多个文件以空格分隔（直接回车则不移除）：')
        if pos:
            pos = pos.split(' ')
            pos.sort(reverse=True)
            for index in pos:
                i = int(index) - 1
                test_excel.pop(i)
        else:
            break
    # print(zhi_pos, zhizhi_excel, pos, test_excel)
    
def disp_dir():
    pos = None
    dir_length = len(dir_path) + 1
    print('=============================================================================')
    for index, file in enumerate(test_excel):
        if '纸' in file:
            pos = index
        print(' %2d : %s' % (index + 1, test_excel[index][dir_length:]))
    print('=============================================================================')
    return pos

def list_dir():
    for name in glob(dir_path + '/*.xls*'):
        test_excel.append(name)

    

if __name__ == '__main__':
    zhizhi_excel = None
    test_excel   = []
    print('=============================================================================')
    print(main.__doc__)
    print('=============================================================================')
    # 包含文件目录参数或参数长度为空（运行批处理时会出现）
    if len(sys.argv) < 2 or not len(sys.argv[1]):
        dir_path = get_dir_path()
    else:
        dir_path = sys.argv[1]
    while True:
        print('数据文件所在位置：{}'.format(dir_path))
        list_dir()
        if len(test_excel) < 2:
            print('数据文件数量小于 2 个或不存在，请重新设置...')
            dir_path = get_dir_path()
        else:
            break
    select_zhizhi()
    main()
    print('=============================================================================')
    print('  设计：王佳辉  bdwjh@163.com   欢迎您的再次使用')
    print('=============================================================================')
    i = input()
