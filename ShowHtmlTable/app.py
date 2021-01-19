# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from api.mysql_func import *
import xlsxwriter
import uuid
import io
import time

app = Flask(__name__)


@app.route('/')
def home():
    # 页面重定向
    userUuid = request.args.get('uuid')
    if check_user(userUuid):
        return redirect(url_for('search', uuid=userUuid))
    else:
        return redirect(url_for('error404'))


@app.route('/404')
def error404():
    return render_template('404.htm')


@app.route('/index')
def index():
    # 每页显示数量 page_show_count
    page_show_count = int(request.args.get('page_show_count')) if request.args.get('page_show_count') != None else 15
    page = int(request.args.get('page')) if request.args.get('page') != None else 1
    # 计算最大页数
    sql_row_count = 'select count(0) from orderInfo where isDel = 0'
    maxPage = (mysql_conn(sql_row_count)[0][0] % page_show_count) / page_show_count if mysql_conn(sql_row_count)[0][
                                                                                           0] % page_show_count == 0 else int(
        mysql_conn(sql_row_count)[0][0] / page_show_count) + 1
    sql = 'select id,shopName,goodsName,goodsKey,wangwangId,orderId,goodsPrice,goodsYj,handlerName,date from orderInfo where isDel = 0 order by id limit %s,%s' % (
        str((page - 1) * page_show_count), page_show_count)
    print('打印sql语句: ', sql)
    sqlRes = mysql_conn(sql)
    lastPage = page - 1 if page > 1 else 1
    nextPage = page + 1 if page < maxPage else maxPage
    print(page, lastPage, nextPage, maxPage)
    return render_template('index.html', sqlRes=sqlRes, lastPage=lastPage, nextPage=nextPage, maxPage=maxPage)


@app.route('/search', methods=["GET", "POST"])
def search():
    # 页面重定向，如果UUID是管理员则显示更多属性，如果没有UUID则跳转404，否则显示普通属性
    # 获取UUID
    userUuid = ''
    role = 0
    formParameters = []
    if request.method == 'GET':
        userUuid = request.args.get('uuid')
    elif request.method == 'POST':
        userUuid = request.form['userUuid']

    if not check_user(userUuid):
        return redirect(url_for('error404'))
    else:
        # 根据UUID获取角色值 9为管理员|0普通用户
        role = int(mysql_conn('select role from userInfo where uuid = {0}'.format('\'' + userUuid + '\''))[0][0])
        # 初始化变量page和pageSize
        if request.method == 'POST':
            form = request.form
            pageSize = int(form['pageSize'])
            page = 1
        elif request.method == 'GET':
            # 如果url中传过来的pageSize为空则每页显示10条数据
            pageSize = int(request.args.get('pageSize')) if request.args.get('pageSize') != None else 10
            # 如果url中传过来的page为空则显示第1页
            page = int(request.args.get('page')) if request.args.get('page') != None else 1
        else:
            page = 1
            pageSize = 10
        # 获取前端传入的模糊查询变量并计算查询结果数量
        if request.method == 'POST':
            form = request.form
            goodsName = form['goodsName']
            goodsKey = form['goodsKey']
            wangwangId = form['wangwangId']
            orderId = form['orderId']
            shopName = form['shopName']
            date = form['date']
            handlerName = form['handlerName']
            custName = form['custName']
            # 将前端查询参数写到数组里
            formParameters.append(goodsName)
            formParameters.append(goodsKey)
            formParameters.append(wangwangId)
            formParameters.append(orderId)
            formParameters.append(shopName)
            formParameters.append(date)
            formParameters.append(handlerName)
            formParameters.append(custName)
            searchSql = ''
            if goodsName != '':
                searchSql = searchSql + ' and goodsName like ' + '\'%' + goodsName + '%\''
            if goodsKey != '':
                searchSql = searchSql + ' and goodsKey like ' + '\'%' + goodsKey + '%\''
            if wangwangId != '':
                searchSql = searchSql + ' and wangwangId like ' + '\'%' + wangwangId + '%\''
            if orderId != '':
                searchSql = searchSql + ' and orderId like ' + '\'%' + orderId + '%\''
            if shopName != '':
                searchSql = searchSql + ' and shopName like ' + '\'%' + shopName + '%\''
            if date != '':
                searchSql = searchSql + ' and date like ' + '\'%' + date + '%\''
            if handlerName != '':
                searchSql = searchSql + ' and handlerName like ' + '\'%' + handlerName + '%\''
            if custName != '':
                searchSql = searchSql + ' and custName like ' + '\'%' + custName + '%\''
            # 本次查询结果总记录数
            queryTotalCntSql = 'select count(0) from orderInfo where isDel = 0 {0}'.format(searchSql)
            queryTotalCnt = int(mysql_conn(queryTotalCntSql)[0][0])
        elif request.method == 'GET':
            searchSql = ''
            queryTotalCnt = 0

        # 计算最大页数
        maxPage = count_max_page(tableName='orderInfo', pageSize=pageSize)
        paginateDict = paginate(page=page, size=pageSize)
        sql = 'select id,shopName,goodsName,goodsKey,wangwangId,orderId,goodsPrice,goodsYj,redPackets,ssyj,handlerName,opWechatId,custName,date from orderInfo where isDel = 0 {0} order by id limit {1},{2}'.format(
            searchSql, paginateDict['offset'], paginateDict['limit'])
        sqlRes = mysql_conn(sql)

        # 计算分页订单总价格
        totalKdjPrice = 0
        totalYjPrice = 0
        totalRedPackets = 0
        totalSsyj = 0
        for tp in sqlRes:
            totalKdjPrice += tp[6]
            totalYjPrice += tp[7]
            totalRedPackets += tp[8]
            totalSsyj += tp[9]
        lastPage = paginateDict['before']
        nextPage = paginateDict['next']
        return render_template(
            'search.html',
            queryTotalCnt=queryTotalCnt,
            sqlRes=sqlRes,
            lastPage=lastPage,
            nextPage=nextPage,
            maxPage=maxPage,
            pageSize=pageSize,
            totalKdjPrice=totalKdjPrice,
            totalYjPrice=totalYjPrice,
            totalRedPackets=totalRedPackets,
            totalSsyj=totalSsyj,
            role=role,
            formParameters=formParameters
        )


@app.route("/search/downloadExcel", methods=["GET"])
def download_excel():
    # 获取url参数信息
    print('/search/downloadExcel',request.args)
    userUuid = request.args.get('uuid')
    goodsName = request.args.get('goodsName')
    goodsKey = request.args.get('goodsKey')
    wangwangId = request.args.get('wangwangId')
    orderId = request.args.get('orderId')
    shopName = request.args.get('shopName')
    date = request.args.get('date')
    handlerName = request.args.get('handlerName')
    custName = request.args.get('custName')
    searchSql = ''
    if goodsName != '':
        searchSql = searchSql + ' and goodsName like ' + '\'%' + goodsName + '%\''
    if goodsKey != '':
        searchSql = searchSql + ' and goodsKey like ' + '\'%' + goodsKey + '%\''
    if wangwangId != '':
        searchSql = searchSql + ' and wangwangId like ' + '\'%' + wangwangId + '%\''
    if orderId != '':
        searchSql = searchSql + ' and orderId like ' + '\'%' + orderId + '%\''
    if shopName != '':
        searchSql = searchSql + ' and shopName like ' + '\'%' + shopName + '%\''
    if date != '':
        searchSql = searchSql + ' and date like ' + '\'%' + date + '%\''
    if handlerName != '':
        searchSql = searchSql + ' and handlerName like ' + '\'%' + handlerName + '%\''
    if custName != '':
        searchSql = searchSql + ' and custName like ' + '\'%' + custName + '%\''

    # 拼接form参数sql
    whereSql = 'select id,shopName,goodsName,goodsKey,wangwangId,orderId,goodsPrice,goodsYj,redPackets,ssyj,handlerName,opWechatId,custName,date from orderInfo where isDel = 0 {0}'.format(searchSql)
    # 获取查询结果
    searchData = mysql_conn(whereSql)

    # 根据UUID查询角色信息，如果是管理员则导出内部可见列，普通角色不导出内部可见列
    header_list = []
    role = int(mysql_conn('select role from userInfo where uuid = {0}'.format('\'' + userUuid + '\''))[0][0])
    if role == 9:
        header_list = ["序号", "店铺名称", "宝贝标题", "关键词", "旺旺", "订单号", "实付金额", "佣金", "红包及其他", "刷手佣金", "经手人", "操作微信号", "客户名称","日期" ]
    else:
        header_list = ["序号", "店铺名称", "宝贝标题", "关键词", "旺旺", "订单号", "实付金额", "佣金", "客户名称", "日期",]
    """1. 生成表头   2. 生成数据  3. 个性化合并单元格，修改字体属性、修改列宽  3. 返回给前端"""
    fp = io.BytesIO()  # 生成一个BytesIO对象
    book = xlsxwriter.Workbook(fp)  # 可以认为创建了一个Excel文件
    worksheet = book.add_worksheet('sheet1')  # 增加一个sheet
    # 1. 生成表头
    header_list = ["序号", "店铺名称", "宝贝标题", "关键词", "旺旺", "订单号", "实付金额", "佣金", "红包及其他", "刷手佣金", "经手人", "操作微信号", "客户名称","日期"]
    for col, header in enumerate(header_list):
        worksheet.write(0, col, header)  # 行(从0开始), 列(从0开始)， 内容

    # 2. 生成数据
    x = 1
    # print('searchData type is',type(searchData))
    # print('searchData value is',searchData)
    for orderInfo in searchData:
        y = 0
        # print('orderInfo values is ',orderInfo)
        for cellInfo in orderInfo:
            # print('cellInfo value is ',cellInfo)
            worksheet.write(x, y, cellInfo)  # 遍历导入每条订单信息
            y += 1
        x += 1

    # 3. 个性化合并单元格，修改字体属性、修改列宽
    # 定义格式实例, 16号字体，加粗，水平居中，垂直居中，红色字体
    my_format = book.add_format(
        {'font_size': 16, 'bold': True, 'align': 'center', 'valign': 'vcenter', "font_color": "red"})
    # worksheet.merge(len(students_data + 1, students_data + 2, 1, 5, "合并单元格内容", my_format))
    book.close()
    fp.seek(0)
    fileName = time.strftime('%Y-%m-%d_%H_%M_%S',time.localtime(time.time()))+'_'+''.join(str(uuid.uuid4()).split('-'))+'.xlsx'
    return send_file(fp, attachment_filename=fileName, as_attachment=True)


if __name__ == '__main__':
    app.run()
