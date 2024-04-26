# -*- coding: utf-8 -*-
from WordFunctions import AddHeadText
from WordFunctions import AddTitle
from WordFunctions import AddParaText
from WordFunctions import document
import xlrd
from xlrd import xldate_as_datetime


def ReadExcel(name):                                   # 获取测试用例excl内容
    excel_data = xlrd.open_workbook(name)               # 打开指定测试用例Excel文件
    name_sheets = excel_data.sheet_names()[0]          # 获取文件内首个表名
    sheet_datas = excel_data.sheet_by_name(name_sheets)          # 遍历获取首表的内容

    return sheet_datas
# 获取测试用例的数据


def createTable(row2, test_case_name2):
    get_sheet_datas2 = ReadExcel(test_case_name2)  # 二次封装调用获取测试用例方法
    table = document.add_table(rows=16, cols=2, style='Light Grid Accent 2')
    w_cells = table.rows[0].cells
    w_cells[0].text = u'用例名称'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 7)  # 用例名称数据输入
    w_cells = table.rows[1].cells
    w_cells[0].text = u'测试预置条件'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 8)  # 测试预置条件输入
    w_cells = table.rows[2].cells
    w_cells[0].text = u'测试业务场景'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 9)  # 测试业务场景输入
    w_cells = table.rows[3].cells
    w_cells[0].text = u'测试数据'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 10)  # 测试数据输入
    w_cells = table.rows[4].cells
    w_cells[0].text = u'测试类型'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 11)  # 测试类型输入
    w_cells = table.rows[5].cells
    w_cells[0].text = u'逻辑步骤'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 12)  # 逻辑步骤输入
    w_cells = table.rows[6].cells
    w_cells[0].text = u'关联系统'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 13)  # 关联系统输入
    w_cells = table.rows[7].cells
    w_cells[0].text = u'涉及文件'
    w_cells[1].text = u'无'
    w_cells = table.rows[8].cells
    w_cells[0].text = u'预期结果'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 15)  # 预期结果输入
    w_cells = table.rows[9].cells
    w_cells[0].text = u'实际结果'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 15)  # 实际结果输入
    w_cells = table.rows[10].cells
    w_cells[0].text = u'附件信息'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 15)  # 附件信息输入
    w_cells = table.rows[11].cells
    w_cells[0].text = u'测试结论'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 17)  # 测试结论输入
    w_cells = table.rows[12].cells
    w_cells[0].text = u'测试时间'
    w_cells[1].text = str(xldate_as_datetime(get_sheet_datas2.cell_value(row2, 20), 0)) + '—' + str(xldate_as_datetime(get_sheet_datas2.cell_value(row2, 21),0))   #  测试结束时间输入
    w_cells = table.rows[13].cells
    w_cells[0].text = u'测试人员'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 16)  # 测试人员输入
    w_cells = table.rows[14].cells
    w_cells[0].text = u'测试报告编写人员'
    w_cells[1].text = get_sheet_datas2.cell_value(row2, 16)  # 测试报告编写人员输入
    w_cells = table.rows[15].cells
    w_cells[0].text = u'报告编写时间'
    w_cells[1].text = str(xldate_as_datetime(get_sheet_datas2.cell_value(row2, 21), 0))  # 报告编写时间输入
    # 创建word测试报告中所需的测试用例表格 #


def CreateWord(test_case_name):
    #   编写测试报告
    get_sheet_datas = ReadExcel(test_case_name) # 二次封装调用获取测试用例方法
    SystemName = get_sheet_datas.cell_value(3, 0)
    #   所属项目
    DemandOrderNo = get_sheet_datas.cell_value(3, 1)
    #   需求编号
    DemandName = get_sheet_datas.cell_value(3, 2)
    #   需求名称
    ParaTestDemand = SystemName + ' ' + DemandOrderNo + ' ' + DemandName
    #   拼接后的“测试需求”章节的段落内容
    TestReport = DemandName + '测试报告.doc'
    #   测试报告名称

    AddTitle('测试报告', 18)                    # 标题
    AddHeadText('1测试需求：', 16, 1)           # 第一章节标题
    AddParaText(ParaTestDemand, 12, 1, 1, 0)    # 第一章节段落
    AddHeadText('2测试报告',  16, 1)            # 第二章节标题

    # 初始化标题编号
    level_1_count = 0
    level_2_count = 0
    level_3_count = 0
    level_4_count = 0

    prev_level_1 = None
    prev_level_2 = None
    prev_level_3 = None
    prev_level_4 = None



    # 从第二行开始遍历测试用例Excel的所有行
    for row in range(2, get_sheet_datas.nrows):
        row_data = get_sheet_datas.row_values(row)      # 获取整行数据
        # 获取所需列的数据
        level_1 = row_data[4]
        level_2 = row_data[5]
        level_3 = row_data[6]
        level_4 = row_data[7]

        # 初始化各级标题
        prev_primary_function = None
        prev_secondary_function = None
        prev_third_function = None
        prev_test_case_name = None


        # 处理一级标题
        if level_1 != prev_level_1:
            level_1_count += 1
            level_2_count = 0
            level_3_count = 0
            level_4_count = 0
            prev_level_1 = level_1

            # 生成一级标题（一级功能）编号
            PrimaryFunction = f"{2}.{level_1_count}.{level_1}"

            # 去重写入Word
            if PrimaryFunction != prev_primary_function:
                AddHeadText(PrimaryFunction, 12, 2)  # 一级功能
                prev_primary_function = PrimaryFunction

        # 处理二级标题
        if level_2 != prev_level_2 or level_1 != prev_level_1:
            level_2_count += 1
            level_3_count = 0
            level_4_count = 0
            prev_level_2 = level_2
            prev_level_3 = None  # 重置上一级标题为 None

            # 生成二级标题（二级功能）编号
            SecondaryFunction = f"{2}.{level_1_count}.{level_2_count}.{level_2}"

            # 去重写入Word
            if SecondaryFunction != prev_secondary_function:
                AddHeadText(SecondaryFunction, 12, 3)  # 二级功能
                prev_secondary_function = SecondaryFunction

        # 处理三级标题
        if level_3 != prev_level_3 or level_2 != prev_level_2 or level_1 != prev_level_1:
            level_3_count += 1
            level_4_count = 0
            prev_level_3 = level_3
            prev_level_4 = None  # 重置上一级标题为 None

            # 生成三级标题（三级功能）编号
            ThirdFunction = f"{2}.{level_1_count}.{level_2_count}.{level_3_count}.{level_3}"

            # 去重写入Word
            if ThirdFunction != prev_third_function:
                AddHeadText(ThirdFunction, 12, 4)  # 三级功能
                prev_third_function = ThirdFunction

        # 处理四级标题
        if level_4 != prev_level_4 or level_3 != prev_level_3 or level_2 != prev_level_2 or level_1 != prev_level_1:
            level_4_count += 1
            prev_level_4 = level_4

            # 生成四级标题（用例名称）编号
            TestCaseName = f"{2}.{level_1_count}.{level_2_count}.{level_3_count}.{level_4_count}.{level_4}"

            # 去重写入Word
            if TestCaseName != prev_test_case_name:
                AddHeadText(TestCaseName, 12, 5)  # 测试用例名称
                createTable(row, test_case_name)  # 生成单个测试用例表格
                prev_test_case_name = TestCaseName






    document.add_page_break()
    document.save(TestReport)

    print(f"测试报告文档已生成并保存为{TestReport}")
    # 创建保存测试报告
