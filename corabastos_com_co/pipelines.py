# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html
from scrapy import signals
import xlsxwriter
import os, xlrd, calendar
from collections import OrderedDict
from datetime import timedelta, datetime

class CorabastosComCoPipeline(object):
    @classmethod
    def from_crawler(cls, crawler):
        pipeline = cls()
        crawler.signals.connect(pipeline.spider_opened, signals.spider_opened)
        crawler.signals.connect(pipeline.spider_closed, signals.spider_closed)
        return pipeline

    def spider_opened(self, spider):

        file_path = 'output/result2.xlsx'
        spider.file_path = file_path
        if not os.path.isfile(file_path):
            print('not existing')

            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet('data')

            years = [2019, 2018, 2017, 2016, 2015, 2014, 2013, 2012, 2011, 2010]
            months = ['12', '11', '10', '09', '08', '07', '06', '05', '04', '03', '02', '01']
            nn = 1
            for i, y in enumerate(years):
                for j, m in enumerate(months):
                    if y == 2019:
                        if j < 2:
                            continue
                    start_date = '{}-{}'.format(str(y), m)
                    worksheet.write(nn, 0, start_date)#write station id on first row
                    nn += 1

            # for i, id_vavl in enumerate(spider.ids):
            #     worksheet.write(0, i + 1, str(id_vavl))#write station id on first row
            workbook.close()
        else:
            book = xlrd.open_workbook(file_path)
            sh = book.sheet_by_index(0)

            spider.now_row_count = sh.nrows
            if spider.now_row_count == 0:
                spider.now_row_count = 1

            # spider.now_row_count += 1

            for row_index in range(sh.nrows):
                if row_index == 0:
                    continue
                a1 = sh.cell_value(rowx=row_index, colx=0)
                if not isinstance(a1, str):
                    a1_as_datetime = datetime(*xlrd.xldate_as_tuple(a1, book.datemode))
                    date_str = a1_as_datetime.strftime("%Y-%m")
                    spider.all_date_data.append(date_str)
                else:
                    spider.all_date_data.append(a1)

            for col_index in range(sh.ncols):
                if col_index == 0:
                    continue
                a1 = sh.cell_value(rowx=0, colx=col_index)
                if isinstance(a1, float) or isinstance(a1, int):
                    a1 = str(int(a1))

                spider.all_fields.append(a1)
            book.release_resources()
            del book
            print('existing')

    def spider_closed(self, spider):
        # spider.workbook_to_write.close()
        spider.wb.save(spider.file_path)
        pass

    def process_item(self, item, spider):
        return item
