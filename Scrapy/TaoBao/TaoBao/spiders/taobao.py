import scrapy
from TaoBao.items import TaobaoItem


class TaobaoSpider(scrapy.Spider):
    name = 'taobao'
    allowed_domains = ['taobao.com']
    start_urls = ['https://s.taobao.com/search?q=手机&s=0']

    def start_requests(self):
        keywords = ['手机', '笔记本电脑']
        for keyword in keywords:
            for page in range(2):
                url = f'https://s.taobao.com/search?q={keyword}&s={page * 44}'
                yield scrapy.Request(url=url)

    # def parse_detail(self, request: scrapy.Request):
    #     pass

    def parse(self, response):
        page_items = response.xpath("//div[@class='item J_MouserOnverReq  ']")
        for page in page_items:
            item = TaobaoItem()
            item['title'] = page.xpath("//div[@class='row row-2 title']/a/span[1]/text()").extract()
            item['price'] = page.xpath("//div[@class='price g_price g_price-highlight']/strong/text()").extract()
            item['deal_count'] = page.xpath("//div[@class='deal-cnt']/text()").extract()
            item['shop'] = page.xpath("//div[@class='shop']/a/span[2]/text()").extract()
            item['location'] = page.xpath("//div[@class='location']/text()").extract()
            yield item
