# -*- coding: utf-8 -*-
# @Time : 2021/1/18 0018 10:35
# @Author : Owen
# @Email : zhangwx0794@gmail.com
# @File : main.py
# @Project : Taobao
from ClsTaobao import *



if __name__ == '__main__':

    xlsList = []
    # 定义xls存放目录
    xlsDir = os.getcwd()+'\\'+'work'+'\\'
    taobao = Taobao()

    # 调用测试方法
    # taobao.testFunc()
    # exit(0)newXlsPath

    xlsList = taobao.get_path_xls(xlsDir)
    xlsxList = taobao.get_path_xlsx(xlsDir)

    for xlsx in xlsxList:
        # taobao.chkRepeOrderInXls(xlsDir+xlsx)
        # taobao.delBlankOrderRow(xlsDir+xlsx)
        print('正在处理',xlsx)
        # cnt = taobao.chkXlsOrderUniq(xlsDir+xlsx)
        # if cnt > 0:
        #     print(cnt)
        if taobao.importData(xlsDir+xlsx) == 0:
            break
        # shopName = re.compile('[\u4e00-\u9fff]+').findall(xlsx)[0]
        # date = re.compile('^\d{4}-\d{2}-\d{2}').findall(xlsx)[0]
        # taobao.completeForm(xlsDir+xlsx,13,shopName)
        # taobao.delBlankOrderRow(xlsDir+xlsx)
        # if taobao.data_format_check(xlsDir+xlsx) >= -1:
        #     continue
        # else:
        #     break
        # 删除重复的xls文件
        # taobao.delRepeName(xlsDir+xlsx,xlsxList)
        # taobao.xls_to_xlsx(xlsDir+xls)
        # taobao.format_xls_name(xlsDir+xlsx)
        # taobao.del_col_from_key(xlsDir+xlsx,'时间',1)
    #     v = taobao.getColValues(xlsDir+xlsx,14)
    #     if v == 'None':
    #         print(xlsx,v)
    #         # taobao.insertColum(xlsDir+xlsx,13)
    #         taobao.writeData2Xls(xlsDir+xlsx,'日期',1,14)
    #
    #     colValues.append(v)
    # newCol = list(set(colValues))
    # print(newCol)

    # 平台
    # 1. 重命名为店铺名称
    # 2. 第一列插入【序号】