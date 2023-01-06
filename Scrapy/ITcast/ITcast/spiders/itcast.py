import scrapy
# from ITcast.items import ItcastItem


class ItcastSpider(scrapy.Spider):
    # 爬虫名，启动爬虫时需要的参数（必须）
    name = 'itcast'
    # 爬取域范围，允许爬虫在这个域名下进行爬取（可选）
    allowed_domains = ['www.itcast.cn']
    # 起始url列表，爬虫执行后第一批请求，将从这个列表里获取
    start_urls = ['https://www.itcast.cn/channel/teacher.shtml']

    def parse(self, response):
        node_list = response.xpath("//div[@class='li_txt']")

        #  用来存储所有的item字段的
        # items = []
        for node in node_list:
            # 创建item字段对象，用于存储信息
            # item = ItcastItem()
            # .extract() 将xpath对象转换为Unicode字符串
            name = node.xpath("./h3/text()").extract()
            title = node.xpath("./h4/text()").extract()
            info = node.xpath("./p/text()").extract()

            # item['name'] = name[0]
            # item['title'] = title[0]
            # item['info'] = info[0]
            #
            # yield item
            # return item
            # items.append(item)

        # return items