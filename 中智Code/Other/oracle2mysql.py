# -*- coding: utf-8 -*-
"""
Created on Tue May 26 08:57:59 2020

@author: Long
"""
# pyinstaller --onefile --nowindowed oracle2mysql.py
# pyinstaller -F oracle2mysql.spec
import time
import datetime
import pymysql
import win32api
import pandas as pd
import pyautogui as pg
import cx_Oracle  # 连接oracle使用
from sqlalchemy.types import Integer, NVARCHAR, Float
from sqlalchemy import create_engine  # 连接mysql使用

# =============================================================================
# 按月导入数据到MYSQL
# =============================================================================
# df文本格式和数据库文本格式转换函数
def mapping_df_types(df):
    dtypedict = {}
    for i, j in zip(df.columns, df.dtypes):
        if "object" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "float" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "int" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
    return dtypedict


def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # 通用写法，起初当月最后一天
    return next_month - datetime.timedelta(days=next_month.day)


def time_0(time):
    if time > 9:
        time1 = str(time)
    else:
        time1 = '0' + str(time)
    return time1


def to_mysql(a, b, c, d, e, f):
    for years in range(a, b):
        for months in range(c, d):
            lasttime = last_day_of_month(datetime.date(years, months, 15))
            year = str(lasttime.year)
            month = time_0(lasttime.month)
            # for day in range(1,lasttime.day+1):
            for day in range(e, f):
                day = time_0(day)
                print("开始导入数据：", year, month, day)
                try:
                    starttime = datetime.datetime.now()

                    # 从oracle数据库里查询对应月份数据，保存到df中
                    # sqlcmd="select * from 表 where 时间 like  " +'\''+ month+'\''
                    # 以下部分为从oracle数据库中读取数据，并导入到mysql中#设置oracle连接参数
                    dsn = cx_Oracle.makedsn("192.168.20.153", 1521, "hydee")  # ip，端口，库名
                    conn = cx_Oracle.connect("zeus_cs", "zeus", dsn, encoding="UTF-8", nencoding="UTF-8")
                    sqlcmd = """SELECT case when s.flag=1 then '线上' else '线下' end as flag_bs,
                                case when s.flag=1 then h.olshopid else to_char(h.busno) end as flag_no,
                                case when s.flag=1 then op.olshopname else s.orgname end as flag_name,
                                omr.order_start_time, vs.vencusno,vs.vencusname,o.olshopid,s.commname,h.MEMBERCARDNO,d.saler,
                                       round(((d.wareqty + (CASE WHEN d.stdtomin = 0 THEN 0 ELSE d.minqty / d.stdtomin END)) * d.times),2) AS wareqty,
                                       d.wareid,t.warecode,t.warename,d.times,
                                       round((CASE
                                         WHEN trunc(h.accdate) < to_date(case when nvl(si.inipara,'0') = '0' then '1990-01-01' else si.inipara end, 'yyyy-mm-dd') THEN
                                          (round((d.stdprice * (d.wareqty + (CASE
                                                   WHEN d.stdtomin = 0 THEN 0 ELSE d.minqty / d.stdtomin
                                                 END)) * d.times),
                                                 2))
                                         ELSE
                                          round((d.stdprice * (d.wareqty + (CASE
                                            WHEN d.stdtomin = 0 THEN 0 ELSE d.minqty / d.stdtomin
                                          END)) * d.times),6)
                                       END),6) AS stdamt,
                                       d.stallno,d.saleno,d.rowno,p.saletax,q.purtax,
                                       nvl(d.purprice, 0) AS purprice,
                                       round((CASE
                                         WHEN trunc(h.accdate) < to_date(case when nvl(si.inipara,'0') = '0' then '1990-01-01' else si.inipara end, 'yyyy-mm-dd') THEN
                                          (round(d.netprice * d.wareqty * d.times +
                                                 d.minqty * d.times * d.minprice,
                                                 2))
                                         ELSE
                                          (round(d.netprice * d.wareqty * d.times,6) +
                                           round(d.minqty * d.times * d.minprice,6))
                                       END),6) AS netamt,
                                       round((CASE
                                         WHEN trunc(h.accdate) < to_date(case when nvl(si.inipara,'0') = '0' then '1990-01-01' else si.inipara end, 'yyyy-mm-dd') THEN
                                          round((d.purprice * (d.wareqty + (CASE
                                                  WHEN d.stdtomin = 0 THEN 0 ELSE d.minqty / d.stdtomin
                                                END)) * d.times),
                                                2)
                                         ELSE
                                          round((d.purprice * (d.wareqty + (CASE
                                                  WHEN d.stdtomin = 0 THEN 0 ELSE d.minqty / d.stdtomin
                                                END)) * d.times),
                                                6)
                                       END),6) AS pursum,
                                       d.wareqty as wareqty_zj,
                                       d.minqty,nvl(d.stdtomin, 0) AS stdtomin,
                                       d.netprice,d.minprice,d.makeno,
                                       d.groupid,d.distype,d.disrate,
                                       d.busno,d.batid,d.accdate,d.invalidate,
                                     (case when p.factoryid=0 then i.factoryid else p.factoryid end) AS factoryid,
                                     (case when p.factoryid=0 then i.factoryid else p.factoryid end) AS factoryid1,
                                       t.warespec,t.wareunit,
                                       round((CASE
                                               WHEN nvl(abs(d.minqty), 0) > 0 AND nvl(abs(stdtomin), 0) > 0 AND
                                                    nvl(abs(stdprice), 0) > 0 THEN minprice * stdtomin / stdprice ELSE (CASE WHEN nvl(abs(stdprice), 0) > 0 THEN netprice / stdprice ELSE 1 END) END), 6) sale_rate,
                                       h.finaltime,
                                       p.waregeneralname,
                                       mr.cardholder
                                  FROM h2.t_sale_h h
                                 INNER JOIN h2.t_sale_d d ON h.saleno = d.saleno
                                 INNER JOIN h2.s_busi s ON h.busno = s.busno and h.compid=s.compid
                                 LEFT JOIN h2.t_store_i i ON d.wareid = i.wareid AND d.batid = i.batid and h.compid=i.compid
                                 INNER JOIN h2.t_ware_base t ON d.wareid = t.wareid
                                 left join h2.t_olshop o on o.olshopid=h.olshopid
                                 left join h2.t_ware p on d.wareid=p.wareid and h.compid=p.compid
                                 left join h2.t_store_i q on d.wareid=q.wareid and d.batid=q.batid and h.compid=q.compid
                                 left join (select a.olpickno,b.order_start_time from
                                (select a.olpickno,max(b.olorderno) as olorderno from h2.t_ol_pick_m a,h2.t_ol_pick_c b where a.olpickno=b.olpickno and a.status=4 group by a.olpickno)
                                a,h2.t_ol_srcorder_m b where a.olorderno=b.olorderno) omr on h.olpickno=omr.olpickno
                                 left join h2.t_vencus vs on i.vencusno=vs.vencusno and i.compid=vs.compid
                                 left join h2.t_olshop op on h.olshopid=op.olshopid
                                 left join h2.t_memcard_reg mr on h.membercardno = mr.memcardno and h.compid = mr.compid and h.busno = mr.busno
                                 left join h2.s_sys_ini si on si.compid = h.compid and si.inicode = '1303'
                                 WHERE 1 = 1
                                 and to_char(h.accdate,'yyyy') = '""" + year + "' and to_char(h.accdate,'mm') = '" + month + "' and to_char(h.accdate,'dd') = '" + day + "'"

                    #                    conn2 = pymysql.connect(
                    #                        host='192.168.249.150',
                    #                        port=3306,
                    #                        user='alex',
                    #                        passwd='123456',
                    #                        db='dkh',
                    #                        charset='utf8'
                    #                    )
                    #                    # 获取游标
                    #                    cursor = conn2.cursor()
                    #                    print("删除旧数据")
                    #                    sql_delect="DELETE FROM v_sale_test WHERE month(FINALTIME) = '"+ month +"' and year(FINALTIME) = '"+ year +"' and day(FINALTIME) = '" + day +"'"
                    #                    cursor.execute(sql_delect)
                    #                    print(cursor.rowcount)
                    #                    conn2.commit()

                    print("提取Oracle数据")
                    df = pd.read_sql(sqlcmd, conn)
                    # 把df列格式转换成数据库格式
                    print("转换数据")
                    dtypedict = mapping_df_types(df)
                    # df保存到mysql
                    # 设置mysql连接引擎
                    print("数据导入MySQ并转好")
                    engine = create_engine('mysql+pymysql://alex:123456@192.168.249.150:3306/123?charset=utf8')
                    df.to_sql('v_sale_test', engine, dtype=dtypedict, index=False, if_exists='append')
                    endtime = datetime.datetime.now()
                    print("导入" + year + month + day + "耗时" + str((endtime - starttime).seconds) + "秒")
                    time.sleep(1)  # 休息10s 个人调试加入，生产可以不需要print("All Finished")
                except Exception as e:
                    print("导入失败：" + year + month + day)
                    print(e)


# 导入会员信息
def HYXX():
    try:
        starttime = datetime.datetime.now()
        # 从oracle数据库里查询对应月份数据，保存到df中
        # sqlcmd="select * from 表 where 时间 like  " +'\''+ month+'\''
        # 以下部分为从oracle数据库中读取数据，并导入到mysql中#设置oracle连接参数
        dsn = cx_Oracle.makedsn("192.168.20.153", 1521, "hydee")  # ip，端口，库名
        conn = cx_Oracle.connect("zeus_cs", "zeus", dsn, encoding="UTF-8", nencoding="UTF-8")
        sqlcmd = """SELECT t1.会员卡号,
                   t1.业务机构代码,
                   t1.会员卡类型,
                   t1.会员卡级别,
                   t1.会员卡状态,
                   t1.会员积分,
                   case when t1.性别 = '男' then 1 when t1.性别 = '女' then 2 else 0  end 性别 ,
                   t1.出生日期,
                   t1.获取途径,
                   t1.申请日期,
                   t1.最后日期,
                   nvl(t1.会员等级编码,0) 会员等级编码,
                   nvl(t2.忠实度编码,0) 忠实度编码,
                   t3.busno chang_busno,
                   t4.countsal
              FROM (SELECT t.memcardno       会员卡号               ,
                           t.busno                        AS 业务机构代码,
                           t.cardtype                     AS 会员卡类型,
                           t.cardlevel                    AS 会员卡级别,
                           t.cardstatus                   AS 会员卡状态,
                           t.integral                     AS 会员积分,
                           t.sex                          AS 性别,
                           t.birthday                     AS 出生日期,
                           t.apptype                      AS 获取途径,
                           t.applytime                    AS 申请日期,
                           t.lastdate                     AS 最后日期,
                           t3.classcode                   AS 类别编码,
                           t_memcard_classgroup.classname AS 类别组,
                           t_memcard_class.classname      AS 会员等级分类,
                           t_memcard_class.classcode      AS 会员等级编码,
                           t.shortchar10           ,
                           t.weixinid
                      FROM h2.t_memcard_reg t
                      LEFT JOIN h2.t_memcard_class_set t3
                        ON t3.memcardno = t.memcardno
                      LEFT JOIN h2.t_memcard_class_base t_memcard_classgroup
                        ON t3.classgroupno = t_memcard_classgroup.classcode
                       AND t_memcard_classgroup.levels = 0
                      LEFT JOIN h2.t_memcard_class_base t_memcard_class
                        ON t3.classcode = t_memcard_class.classcode
                       AND t3.classgroupno = t_memcard_class.classgroupno
                       AND t_memcard_class.levels > 0
                     WHERE t3.classgroupno IN ('20')
                     ORDER BY t.memcardno) t1
              LEFT JOIN
             (
              SELECT t.memcardno                    AS 会员卡号,
                      t.busno                        AS 业务机构代码,
                      t.cardtype                     AS 会员卡类型,
                      t.cardlevel                    AS 会员卡级别,
                      t.cardstatus                   AS 会员卡状态,
                      t.integral                     AS 会员积分,
                      t.sex                          AS 性别,
                      t.birthday                     AS 出生日期,
                      t.apptype                      AS 获取途径,
                      t.applytime                    AS 申请日期,
                      t.lastdate                     AS 最后日期,
                      t4.classcode                   AS 类别编码,
                      t_memcard_classgroup.classname AS 类别组,
                      t_memcard_class.classcode      AS 忠实度编码,
                      t.shortchar10    ,
                      t.weixinid
                FROM h2.t_memcard_reg t
                LEFT JOIN h2.t_memcard_class_set t4
                  ON t4.memcardno = t.memcardno
                LEFT JOIN h2.t_memcard_class_base t_memcard_classgroup
                  ON t4.classgroupno = t_memcard_classgroup.classcode
                 AND t_memcard_classgroup.levels = 0
                LEFT JOIN h2.t_memcard_class_base t_memcard_class
                  ON t4.classcode = t_memcard_class.classcode
                 AND t4.classgroupno = t_memcard_class.classgroupno
                 AND t_memcard_class.levels > 0
               WHERE t4.classgroupno IN ('02')
               ORDER BY t.memcardno) t2
                ON t2.会员卡号 = t1.会员卡号
            
                left join (
                     select membercardno,busno from
                     (
                     select
                     t_sale_h.*,row_number() over(partition by t_sale_h.membercardno order by 1) rn
                     from h2.t_sale_h
                     order by membercardno,accdate desc
                     ) where rn = 10
                ) t3 on t1.会员卡号 = t3.membercardno
                left join (
                  select membercardno,count(saleno) countsal from
                     (
                     select
                     t_sale_h.busno,t_sale_h.membercardno,row_number() over(partition by t_sale_h.membercardno order by 1) rn,
                     saleno,
                     applytime,
                      to_char(t_memcard_reg.applytime,'yyyy-MM-dd'),
                     to_char(t_memcard_reg.applytime+90,'yyyy-MM-dd'),
                     t_sale_h.accdate
                     from h2.t_sale_h
                     left join h2.t_memcard_reg on t_memcard_reg.MEMCARDNO = t_sale_h.membercardno
                     order by membercardno,accdate desc
                     ) t3
                     left join h2.t_memcard_reg on t_memcard_reg.MEMCARDNO = t3.membercardno
                     WHERE
                      to_char(t3.applytime+90,'yyyy-mm-dd')>to_char(t3.accdate,'yyyy-mm-dd')
                     and to_char(t3.accdate,'yyyy-mm-dd')>to_char(t3.applytime,'yyyy-mm-dd')
                     group by membercardno
            
                ) t4 on t1.会员卡号 = t4.membercardno
                where  t1.会员卡号 not like '% %'
                """

        #        conn2 = pymysql.connect(
        #            host='192.168.249.150',
        #            port=3306,
        #            user='alex',
        #            passwd='123456',
        #            db='dkh',
        #            charset='utf8'
        #        )
        #        # 获取游标
        #        cursor = conn2.cursor()
        #        print("删除旧数据")
        #        sql_delect="DELETE FROM v_sale_test WHERE month(FINALTIME) = '"+"' and year(FINALTIME) = '"+"' and day(FINALTIME) = '" +"'"
        #        cursor.execute(sql_delect)
        #        print(cursor.rowcount)
        #        conn2.commit()
        print("提取Oracle数据")
        df = pd.read_sql(sqlcmd, conn)
        # 把df列格式转换成数据库格式
        print("转换数据")
        dtypedict = mapping_df_types(df)
        # df保存到mysql
        # 设置mysql连接引擎
        print("数据导入MySQL")
        engine = create_engine('mysql+pymysql://alex:123456@192.168.249.150:3306/123?charset=utf8')
        df.to_sql('v_dsj_hyxxandhyfl', engine, dtype=dtypedict, index=False, if_exists='replace')
        endtime = datetime.datetime.now()
        print("导入耗时" + str((endtime - starttime).seconds) + "秒")
        time.sleep(1)  # 休息10s 个人调试加入，生产可以不需要print("All Finished")
    except Exception as e:
        print("导入失败：", e)


if __name__ == '__main__':
    conn2 = pymysql.connect(
        host='192.168.249.150',
        port=3306,
        user='alex',
        passwd='123456',
        db='123',
        charset='utf8'
    )
    # 获取游标
    print("查看日期")
    sqlcmd = "SELECT MAX(FINALTIME) FROM v_sale_test"
    # 利用pandas 模块导入mysql数据
    a = pd.read_sql(sqlcmd, conn2)  # 取数据库最大一天

    a1 = a["MAX(FINALTIME)"][0]
    c = time.strptime(str(a1), "%Y-%m-%d %H:%M:%S")  # 转换格式
    b = time.localtime()  # 取今天
    # 结构化时间转换为时间戳格式
    struct_time1, struct_time2 = time.mktime(c), time.mktime(b)  # 转换格式
    # 差的时间戳
    diff_time = struct_time2 - struct_time1  #
    # 将计算出来的时间戳转换为结构化时间
    struct_time = time.gmtime(diff_time)
    # 减去时间戳最开始的时间 并格式化输出
    print('过去了{0}年{1}月{2}日{3}小时{4}分钟{5}秒'.format(
        struct_time.tm_year - 1970,
        struct_time.tm_mon - 1,
        struct_time.tm_mday - 1,
        struct_time.tm_hour,
        struct_time.tm_min,
        struct_time.tm_sec
    ))
    conn2.close

    if struct_time.tm_mday - 1 == 0:
        pass
    else:
        for i in range(1, struct_time.tm_mday):
            # for i in range(1,30):
            nowyear = int((datetime.datetime.now() + datetime.timedelta(days=-i)).strftime("%Y"))
            # nowyear = int(time.strftime("%Y", time.localtime()))  # 年
            nowmonth = int((datetime.datetime.now() + datetime.timedelta(days=-i)).strftime("%m"))  # 月
            nowday = int((datetime.datetime.now() + datetime.timedelta(days=-i)).strftime("%d"))  # 日
            print(nowyear, nowmonth, nowday)
            to_mysql(nowyear, nowyear + 1, nowmonth, nowmonth + 1, nowday, nowday + 1)  # 开始上传昨天的数据

    #    to_mysql(nowyear,nowyear+1,nowmonth,nowmonth+1,nowday-2,nowday-1)#开始上传前2天的数据 2号可能会出错
    #    to_mysql(nowyear,nowyear+1,nowmonth,nowmonth+1,nowday-1,nowday)#开始上传前1天的数据
    #    to_mysql(nowyear,nowyear+1,nowmonth,nowmonth+1,nowday,nowday+1)#开始上传昨天的数据
    HYXX()  # 会员信心数据刷新

    #    to_mysql(2020,2021,4,5,1,32)#选择上传日期范围
    #   import os

    # os.open("D:\\文控中心\\BI中心\\OneDrive\\powerbi\\中智连锁公司BI看板.pbix",os.O_RDWR)

    win32api.ShellExecute(0, 'open', "G:\\中智连锁公司BI看板.pbix", '', '', 1)
    time.sleep(300)  # 等待打开PBI时间
    pg.click(931, 94)  # 点击刷新时间
    time.sleep(5400)  # 等待刷新时间
    pg.click(1581, 19)  # 退出
    time.sleep(1000)  # 等待保存对话框
    pg.click(821, 470)  # 点击保存
    input("please input any key to exit!")
