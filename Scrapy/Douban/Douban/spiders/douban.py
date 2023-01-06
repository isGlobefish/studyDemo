import scrapy
from scrapy import Request, Selector
from Douban.items import MovieItem


class DoubanSpider(scrapy.Spider):
    name = 'douban'
    allowed_domains = ['movie.douban.com']
    start_urls = ['https://movie.douban.com/chart']

    # 翻页先定义好全部的网页地址
    # def start_requests(self):
    #     for page in range(10):
    # 代理方法一：在request此处添加
    #         yield scrapy.Request(url=f'https://movie.douban.com/top250?start={page * 25}&filter=',
    #                              meta={'proxy': 'socks5://127.0.0.1:1086'}
    #                              )

    def parse(self, response, **kwargs):
        # sel = scrapy.Selector(response)
        # .extract() 返回对象
        # .extract_first() 返回列表
        list_items = response.xpath("//div[@class='pl2']")

        for item in list_items:
            # detail_url = item.css('xxxxx').extract_first()
            movie_item = MovieItem()
            movie_item['title'] = str(item.xpath("./a/text()").extract() + item.xpath("./a/span/text()").extract()).replace('\n', '')
            movie_item['content'] = str(item.xpath("./p/text()").extract()).replace('\n', '')
            movie_item['rating_nums'] = str(item.xpath("./div/span[@class='rating_nums']/text()").extract()).replace('\n', '')
            movie_item['people_nums'] = str(item.xpath("./div/span[@class='pl']/text()").extract()).replace('\n', '')
            yield movie_item
            # 每一个电影的详情页信息,高阶回调函数，只传函数名
            # yield Request(url=detail_url, callback=self.detail_parse, cb_kwargs={'item': movie_item})

        # 获取全部页的url
        # url_list = scrapy.xpath("xxxxxxxxxx")
        # for url in url_list:
        #     yield scrapy.Request(url)

    # def detail_parse(self, response, **kwargs):
    #     movie_item = kwargs['item']
    #     sel = Selector(response)
    #     movie_item['duration'] = sel.css('xxxxxxxxxxx').extract()
    #     movie_item['intro'] = sel.css('xxxxxxxxxxx').extract_first()
    #     yield movie_item
