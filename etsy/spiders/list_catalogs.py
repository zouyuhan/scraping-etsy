# -*- coding: utf-8 -*-

import scrapy
import os
import sys
import csv
import glob
import json
from openpyxl import Workbook
from scrapy.http import Request
from etsy.items import ProductItem
from scrapy.loader import ItemLoader

from .product_info import ProductDetailsSpider

# Spider Class
class CatalogsSpider(scrapy.Spider):
    # Spider name
    name = 'list_catalogs'
    allowed_domains = ['etsy.com']
    start_urls = ['https://www.etsy.com/']

    # Get only the products URLs
    URLS_ONLY = False

    product_details_spider = None

    def __init__(self, catalogs, reviews_option=1, count_max=None, urls_only=False, *args, **kwargs):
        if catalogs:
            # Build the search URL
            self.start_urls = [f'https://www.etsy.com/hk-en/c/{catalogs}?page=1']

            # Get only the products URLs
            self.URLS_ONLY = bool(urls_only)

        print(f"#### start_urls: {self.start_urls}")

        self.product_details_spider = ProductDetailsSpider(
                reviews_option, 
                count_max,
                *args,
                **kwargs) 

        super(CatalogsSpider, self).__init__(*args, **kwargs)


    # Parse the first page result and go to the next page
    def parse(self, response):
        print(f"#### response: {response}")

        # Get the list of products from html response
        products_link_list = response.xpath('//div[@data-search-results=""]/div//ol//li//a[1]/@href').extract()

        # For each product extracts the product URL
        print(f"#### FOUND {len(products_link_list)} PRODUCTS:")

        if self.URLS_ONLY:
            for product_link in products_link_list:

                # Create the ItemLoader object that stores each product information
                l = ItemLoader(item=ProductItem(), response=response)

                l.add_value('url', product_link)
                yield l.load_item()

        else:
            for product_link in products_link_list:
                try:
                    # ex: https://www.etsy.com/search or https://www.etsy.com/hk-en/search
                    if product_link.split('/')[3].startswith('search')  \
                            or product_link.split('/')[4].startswith('search'):
                        continue
                except IndexError:
                    continue

                # Go to the product's page to get the data
                yield scrapy.Request(
                        product_link, 
                        callback=self.product_details_spider.parse_product, 
                        dont_filter=True)

        # Pagination - Go to the next page
        current_page_number = int(response.url.split('=')[-1])
        next_page_number = current_page_number + 1
        # Build the next page URL
        next_page_url = '='.join(response.url.split('=')[:-1]) + '=' + str(next_page_number)

        # If the current list is not empty
        if len(products_link_list) > 0:
            yield scrapy.Request(next_page_url)


    # Create the Excel file
    def close(self, reason):
        # Check if there is a CSV file in arguments
        csv_found = False
        for arg in sys.argv:
            if '.csv' in arg:
                csv_found = True

        if csv_found:
            self.logger.info('Creating Excel file')
            #  Get the last csv file created
            csv_file = max(glob.iglob('*.csv'), key=os.path.getctime)

            wb = Workbook()
            ws = wb.active

            with open(csv_file, 'r', encoding='utf-8') as f:
                for row in csv.reader(f):
                    # Check if the row is not empty
                    if row:
                        ws.append(row)
            # Saves the file
            wb.save(csv_file.replace('.csv', '') + '.xlsx')
