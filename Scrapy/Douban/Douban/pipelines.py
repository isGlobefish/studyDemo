# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


# useful for handling different item types with a single interface
from itemadapter import ItemAdapter

import openpyxl
import pymysql


class ExcelPipeline:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = "Top250"
        # 子页追加一行数据
        self.ws.append(('标题', '内容', '评分', '人数'))

    def open_spider(self, spider):
        pass

    def process_item(self, item, spider):
        title = item.get('title', '')
        content = item.get('content', '')
        rating_nums = item.get('rating_nums', '')
        people_nums = item.get('people_nums', '')
        self.ws.append((title, content, rating_nums, people_nums))
        return item

    def close_spider(self, spider):
        self.wb.save('电影数据.xlsx')
        self.wb.close()


class DBPipeline:
    def __init__(self):
        self.conn = pymysql.connect(host='localhost',
                                    port=3306,
                                    user='root',
                                    password='88888888',
                                    database='my_test',
                                    charset='utf8mb4')
        self.cursor = self.conn.cursor()

    def open_spider(self, spider):
        pass

    def process_item(self, item, spider):
        self.cursor.execute("""
         select
            B.Name
            from
            (
            select
            A.Name,
            A.Date,
            date_sub(A.Date,interval A.rn day)  AS inteval_days
            from
            (
                select
                Name,
                Date,
                row_number() over (partition by Name order by Date) AS rn
                from
                day20
            )A
            )B
            group by B.Name,B.inteval_days
            having count(1) >= 3;
        """)
        self.conn.commit()
        self.result = self.cursor.fetchall()
        print('', self.result)
        return item

    def close_spider(self, spider):
        self.cursor.close()
        self.conn.close()
