# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html

# This Pipeline processes several items scraped.
class EtsyPipeline(object):

    def process_item(self, item, spider):
        print(f"#### DEBUG price: {item['price']}")
        print(f"#### DEBUG sales: {item['sales']}")

        # Format the price output
        if 'price' in item:
            item['price'] = item['price'].split(' ')[-1]
            # Remove the currency symbol and the + signal
            item['price'] = item['price'].replace('$', '').replace('+', '')

        if 'sales' in item:
            item['sales'] = item['sales'].split(' ')[0]
            item['sales'] = item['sales'].replace(',', '')

        # Remove the 'in' string 
        if 'store_location' in item:
            item['store_location'] = item['store_location'].replace('in ', '')
        
        # Remove the 'From' string 
        if 'return_location' in item:
            item['return_location'] = item['return_location'].replace('From ', '')

        # Sometimes the spider take the rate in the wrong format (ex: 48.333 instead of 4.8333)
        if 'rating' in item:
            rating = item['rating']
            if float(rating) > 5:
                # Ex: Transform this 48.333 in 4.83)
                rating = rating[0] + '.' + rating[1:2]
            else:
                rating = round(float(rating), 2)

            item['rating'] = rating
        
        return item     
