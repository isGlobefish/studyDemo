# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


# useful for handling different item types with a single interface
from itemadapter import ItemAdapter

import pymysql
from scrapy.crawler import Crawler


class TaobaoPipeline:

    @classmethod
    def from_crawler(cls, crawler: Crawler):
        host = crawler.settings['DB_HOST']
        port = crawler.settings['DB_PORT']
        username = crawler.settings['DB_USER']
        password = crawler.settings['DB_PASSWORD']
        database = crawler.settings['DB_DATABASE']
        charset = crawler.settings['DB_CHARSET']
        return cls(host, port, username, password, database, charset)

    def __init__(self, host, port, username, password, database, charset):
        self.conn = pymysql.connect(host=host,
                                    port=port,
                                    user=username,
                                    password=password,
                                    database=database,
                                    charset=charset,
                                    autocommit=True)
        self.cursor = self.conn.cursor()

    def open_spider(self, spider):
        pass

    def process_item(self, item, spider):
        title = item.get('title', '')
        price = item.get('price', '')
        deal_count = item.get('deal_count', '')
        shop = item.get('shop', '')
        location = item.get('location', '')
        self.cursor.execute(
            "insert into taobao_goods (title, price, deal_count, shop, location)VALUES ({a},{b},{c},{d},{e})".format(a=title, b=price, c=deal_count, d=shop, e=location)
        )
        self.conn.commit()
        return item

    def close_spider(self, spider):
        self.cursor.close()
        self.conn.close()
