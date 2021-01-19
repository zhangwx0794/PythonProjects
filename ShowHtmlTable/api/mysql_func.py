# -*- coding: utf-8 -*-
import pymysql

def check_user(userUuid):
    if userUuid == None:
        return False
    else:
        userUuid = '\''+str(userUuid)+'\''
        sql = 'select count(0) from userInfo where isDel = 0 and uuid = %s' % userUuid
        sqlCnt = int(mysql_conn(sql)[0][0])
    return True if sqlCnt == 1 else False

def mysql_conn(sql):
    conn = pymysql.Connect(
        host='47.111.113.238',
        port=3318,
        db='taobao',
        user='taobao',
        passwd='Boss123456..taobao'
    )
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
    except Exception as e:
        print(sql)
        print('数据库事务执行异常! ',e)
        conn.rollback()
    else:
        # print('数据库事务执行成功!')
        conn.commit()
    res = cursor.fetchall()
    # 关闭数据库
    cursor.close()
    conn.close()
    return res

def count_max_page(tableName,pageSize):
    pageSize = int(pageSize)
    sql = 'select count(0) from {0} where isDel = 0'.format(tableName)
    sqlCount = int(mysql_conn(sql)[0][0])
    if sqlCount % pageSize == 0 and sqlCount != 0:
        maxPage = sqlCount / pageSize
    else:
        maxPage = int(sqlCount / pageSize) + 1
    return maxPage

def paginate(page, size=20):
    """
    数据库 分页 和 翻页 功能函数
    @param page: int or str 页面页数
    @param size: int or str 分页大小
    @return: dict
    {
        'limit': 20,   所取数据行数
        'offset': 0,   跳过的行数
        'before': 0,   前一页页码
        'current': 1,  当前页页码
        'next': 2      后一页页码
    }
    """

    if not isinstance(page, int):
        try:
            page = int(page)
        except TypeError:
            page = 1
    if not isinstance(size, int):
        try:
            size = int(size)
        except TypeError:
            size = 20
    if page > 0:
        page -= 1

    data = {
        "limit": size,
        "offset": page * size,
        "before": page,
        "current": page + 1,
        "next": page + 2
    }
    return data