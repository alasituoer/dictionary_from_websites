#coding:utf-8
import scrapy
import time
from openpyxl import Workbook, load_workbook


RENDER_HTML_URL = "http://localhost:8050/render.html"

class TaobaoSpider(scrapy.Spider):
    name = 'taobao_spider'
    start_urls = ['https://www.taobao.com']
    allowed_domains = ['taobao.com', 'localhost',]

    path_to_write = 'data/dict_from_taobao_' +\
            time.strftime("%Y%m%d", time.localtime()) + '.xlsx'

    def __init__(self):
        try:
            wb = load_workbook(self.path_to_write)
        except Exception, e:
            wb = Workbook()
	    wb.save(self.path_to_write)

    def parse(self, response):
        for sel in response.xpath("//ul[@class='service-bd']"):
            list_cate_name = sel.xpath("li/a/text()").extract()
            list_cate_url = sel.xpath("li/a/@href").extract()
        # (汽车)用品 VS (母婴)用品, 两个用品均可去掉
        dict_cate_url = dict(zip(list_cate_name, list_cate_url))
        dict_cate_url.pop(u'用品')
#        print len(list_cate_name)
#        print len(dict_cate_url.keys())

        list_cate = [
                u'女装', u'男装', u'内衣', u'鞋靴', u'箱包',]# u'配件',
#                u'童装玩具', u'孕产', u'家电', u'数码', u'手机',
#                u'美妆', u'洗护', u'保健品', u'珠宝', u'眼镜', u'手表',
#                u'运动', u'户外', u'乐器', u'游戏', u'动漫', u'影视',
#                u'美食', u'生鲜', u'零食', u'鲜花', u'宠物', u'农资',
#                u'房产', u'装修', u'家具', u'家饰', u'家纺', u'汽车',
#                u'二手车', u'办公', u'DIY', u'五金电子', u'百货', u'货厨',
#                u'家庭保健', u'学习', u'卡券', u'本地服务',]
        list_method = [self.crawlingNvzhuang, self.crawlingNanzhuang,
                self.crawlingNeiyi, self.crawlingXie, self.crawlingXiangbao]
        dict_cate_method = dict(zip(list_cate, list_method))

        for cate_name in list_cate[-1:]:
            cate_url = dict_cate_url[cate_name]
#            print cate_name, cate_url
#            cate_url = RENDER_HTML_URL + "?url=" + cate_url + "&timeout=10&wait=2"
            yield scrapy.Request(url=cate_url, callback=dict_cate_method[cate_name])

    def crawlingXiangbao(self, response):
        """箱包"""
        wb = load_workbook(self.path_to_write)
        try:
            wb.remove_sheet(wb[u'箱包'])
        except Exception, e:
            pass
        ws = wb.create_sheet(title=u'箱包')

        for i,sel in enumerate(response.xpath("//dl[@class='theme-bd-level2']")[:1]):
            list_old_cate = eval(sel.xpath("textarea[1]/text()").extract()[0].strip())
            list_old_cate = [c['cat_name'].decode('utf-8') for c in list_old_cate]
#            print repr(list_old_cate).decode('unicode-escape')
            for kw in list_old_cate[1:]:
#                print list_old_cate[0], kw
                ws.append([2*i+1, list_old_cate[0], kw])

            list_extra_cate = eval(sel.xpath("textarea[2]/text()").extract()[0].strip())
            list_cate = [d['cat_name'].decode('utf-8') for d in list_extra_cate]
	    list_istitle = [d['is_title']=='true' for d in list_extra_cate]
            print repr(list_cate).decode('unicode-escape')
	    print list_istitle


#            for kw in list_extra_cate[1:]:
#                print list_extra_cate[0], kw
#                ws.append([2*i+2, list_extra_cate[0], kw])

        """
        wb.save(self.path_to_write)
	"""


    def crawlingXie(self, response):
        """鞋靴"""
        wb = load_workbook(self.path_to_write)
        try:
            wb.remove_sheet(wb[u'鞋靴'])
        except Exception, e:
            pass
        ws = wb.create_sheet(title=u'鞋靴')

        for i,sel in enumerate(response.xpath("//dl[@class='theme-bd-level2']")):
            list_old_cate = eval(sel.xpath("textarea[1]/text()").extract()[0].strip())
            list_old_cate = [c['cat_name'].decode('utf-8') for c in list_old_cate]
#            print repr(list_old_cate).decode('unicode-escape')
            for kw in list_old_cate[1:]:
#                print list_old_cate[0], kw
                ws.append([2*i+1, list_old_cate[0], kw])

            # 每一大类的扩展类单独作为一个大类
            list_extra_cate = eval(sel.xpath("textarea[2]/text()").extract()[0].strip())
            list_extra_cate = [c['cat_name'].decode('utf-8') for c in list_extra_cate]
#            print repr(list_extra_cate).decode('unicode-escape')
            for kw in list_extra_cate[1:]:
#                print list_extra_cate[0], kw
                ws.append([2*i+2, list_extra_cate[0], kw])

        wb.save(self.path_to_write)


    def crawlingNeiyi(self, response):
        """内衣"""
	print 'hello alas'

        with open('data/t.html', 'w') as f:
	    f.write(response.body)

	"""
        wb = load_workbook(self.path_to_write)
        try:
            wb.remove_sheet(wb[u'内衣'])
        except Exception, e:
            pass
        ws = wb.create_sheet(title=u'内衣')

        for i,sel in enumerate(response.xpath("//ul[@class='list-wrap']/li")):
            c1 = sel.xpath("p/a/text()").extract()
            list_kws_c1 = sel.xpath("dl/dd/a/text()").extract()
            print c1, repr(list_kws_c1).decode('unicode-escape')
            for kw in list_kws_c1:
                list_to_write = [i+1, c1, kw]
                print repr(list_to_write).decode('unicode-escape')
                ws.append(list_to_write)
        wb.save(self.path_to_write)
	"""


    def crawlingNanzhuang(self, response):
        """男装"""
        wb = load_workbook(self.path_to_write)
        try:
            wb.remove_sheet(wb[u'男装'])
        except Exception, e:
            pass
        ws = wb.create_sheet(title=u'男装')

        for i,sel in enumerate(response.xpath("//dl[@class='theme-bd-level2']")):
            c1 = sel.xpath("dt/div/a/text()").extract()[0]
            list_kws_c1 = sel.xpath("dd/a/text()").extract()
#            print c1, repr(list_kws_c1).decode('unicode-escape')
            for kw in list_kws_c1:
                list_to_write = [i+1, c1, kw]
#                print repr(list_to_write).decode('unicode-escape')
                ws.append(list_to_write)
        wb.save(self.path_to_write)


    def crawlingNvzhuang(self, response):
        """女装"""
        wb = load_workbook(self.path_to_write)
        try:
            wb.remove_sheet(wb[u'女装'])
        except Exception, e:
            pass
        ws = wb.create_sheet(title=u'女装')

        for i,sel in enumerate(response.xpath("//ul[@class='list-wrap']/li")):
            c1 = sel.xpath("p/a/text()").extract()[0]
            list_kws_c1 = sel.xpath("dl/dd/a/text()").extract()
#            print c1, repr(list_kws_c1).decode('unicode-escape')
            for kw in list_kws_c1:
                list_to_write = [i+1, c1, kw]
#                print repr(list_to_write).decode('unicode-escape')
                ws.append(list_to_write)
        wb.save(self.path_to_write)



