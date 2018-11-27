# -*- coding: utf-8 -*-
"""
Created on Wed Dec 07 15:49:40 2016

@author: agnisha.singh
"""
import urllib2
from random import randint
import traceback
import re
import xlsxwriter
import lxml.html as LH
import time
from bs4 import BeautifulSoup
from os.path import expanduser
home = expanduser("~")
product_list=[]
seller_list=[]
price_list=[]
list_price_list=[]
reviews_list=[]
ratings_list=[]
model_list=[]
urls=["https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_1?_encoding=UTF8&pg=1",
      "https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_2?_encoding=UTF8&pg=2",
      "https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_3?_encoding=UTF8&pg=3"]
#urls=["https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_1/157-7422272-5893109?_encoding=UTF8&pg=1",
#      "https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_2/157-7422272-5893109?_encoding=UTF8&pg=2",
#      "https://www.amazon.com/Best-Sellers-Toys-Games/zgbs/toys-and-games/ref=zg_bs_pg_3/157-7422272-5893109?_encoding=UTF8&pg=3"]
user_agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64)'
headers = {'User-Agent': user_agent}
address_list=[]    
for i in range(3):
    req=urllib2.Request(urls[i],None, headers)        
    soup = BeautifulSoup(urllib2.urlopen(req).read())
    time.sleep(randint(1,10))
    root=LH.fromstring(str(soup))
    total_elements=root.xpath('//div[@class="zg_itemWrapper"]/div/a[@class="a-link-normal"]')    
    for i in range(len(total_elements)):        
        if len(address_list)<50:
            a="https://www.amazon.com" + total_elements[i].attrib['href']
#            wait=input("111") 
            address_list.append(a)
        else:
            break
print(address_list) 
count=1        
for i in range(len(address_list)):
    print(count)
    req=urllib2.Request(address_list[i],None, headers)
    time.sleep(randint(1,10))        
    soup = BeautifulSoup(urllib2.urlopen(req).read()) 
    try:      
        product_name_ele=soup.find('span',{'id':'title'})
        product_list.append(product_name_ele.text)
    except:
        product_name_ele=soup.find('span',{'id':'productTitle'})
        product_list.append(product_name_ele.text)
    try:    
        seller_ele=soup.find('a',{'id':'bylineContributor'})       #check
        seller_list.append(seller_ele.text)
    except:
        seller_ele=soup.find('a',{'id':'brand'})                   #check
        seller_list.append(seller_ele.text.lstrip().strip())
    price=0    
    try:    
        price_element=soup.find("span",{'id':'priceblock_ourprice'})
        price_list.append(price_element.text)
        price=price_element.text
    except:
        try:
            price_element=soup.find("span",{'id':'priceblock_dealprice'})          #check
            price_list.append(price_element.text)
            price=price_element.text  
        except:    
            price="-"
            price_list.append("-")
    try:
        list_price_element=soup.find("span",{'class':'a-text-strike'})          
        list_price_list.append(list_price_element.text)
    except:
        list_price_list.append(price)
    try:    
        reviews_element=soup.find("span",{"data-hook":"total-review-count"})  
        head,sep,tail=reviews_element.text.partition("customer")
        reviews_list.append(head.strip())
    except: 
      try:
        details_table_element=soup.find("table",{"id":"productDetails_detailBullets_sections1"}) 
        if "customer reviews" in details_table_element.text:
            head,sep,tail=details_table_element.text.partition("customer reviews") 
            head,sep,tail=tail.partition("stars")
            head,sep,tail=tail.partition("customer")
            reviews_list.append(head.strip())
      except:
          reviews_list.append("-")
            
    try:    
        ratings_element=soup.find("span",{"class":"averageStarRatingText"})
        head,sep,tail=ratings_element.text.partition("out")
        ratings_list.append(head.strip())
    except:
        try:
            ratings_element=soup.find("a",{"id":"reviewStarsLinkedCustomerReviews"})
            head,sep,tail=ratings_element.text.partition("out")
            ratings_list.append(head.strip())
        except:
            ratings_list.append("-")
    try:
        if "Item model number" in details_table_element.text:
            head,sep,tail=details_table_element.text.partition("Item model number") 
            list_details=tail.split()
            model_list.append(list_details[0])              
        else:
            model_list.append("-")
    except:
       model_list.append("-") 
    count+=1   
      
try:
    print(len(product_list))
    print(seller_list)
    print(len(seller_list))
    print(price_list)
    print(len(price_list))
    print(list_price_list)
    print(len(list_price_list))
    print(model_list)
    print(len(model_list))
    print(reviews_list)
    print(len(reviews_list))
    print(ratings_list)
    print(len(ratings_list))
    product_list=[unicode(product) for product in product_list]     
    data_file=xlsxwriter.Workbook("{}/Desktop/Amazon0.xlsx".format(home))
    data_worksheet =data_file.add_worksheet("Action & Toy Figures")
    data_headers=("Product", "Price","List Price","Model", "Seller", "Rating", "Reviews")
    format = data_file.add_format({'bold': True})
    row=0
    col=0
    for header in data_headers:
        data_worksheet.write(row, col, header,format)        
        col=col+1
    for i in range(50):
            data_worksheet.write(i+1,0,product_list[i])
            data_worksheet.write(i+1,1, price_list[i])
            data_worksheet.write(i+1,2, list_price_list[i])
            data_worksheet.write(i+1,3, model_list[i])  
            data_worksheet.write(i+1,4, seller_list[i])
            data_worksheet.write(i+1,5, ratings_list[i])
            data_worksheet.write(i+1,6, reviews_list[i])
    data_file.close()           
except Exception as error:
    traceback.print_exc() 
    print(product_list)
    print(len(product_list))
    print(seller_list)
    print(len(seller_list))
    print(price_list)
    print(len(price_list))
    print(list_price_list)
    print(len(list_price_list))
    print(model_list)
    print(len(model_list))
    print(reviews_list)
    print(len(reviews_list))
    print(ratings_list)
    print(len(ratings_list))
    product_list=[unicode(product) for product in product_list] 
    data_file=xlsxwriter.Workbook("{}/Desktop/Amazon-Exception0.xlsx".format(home))
    data_worksheet =data_file.add_worksheet("Action & Toy Figures")
    data_headers=("Product", "Price","List Price","Model", "Seller", "Rating", "Reviews")
    format = data_file.add_format({'bold': True})
    row=0
    col=0
    for header in data_headers:
        data_worksheet.write(row, col, header,format)        
        col=col+1
    for i in range(len(product_list)):
            data_worksheet.write(i+1,0,product_list[i])
    for i in range(len(price_list)):        
            data_worksheet.write(i+1,1, price_list[i])
    for i in range(len(list_price_list)):    
            data_worksheet.write(i+1,2, list_price_list[i])
    for i in range(len(model_list)):        
            data_worksheet.write(i+1,3, model_list[i]) 
    for i in range(len(seller_list)):         
            data_worksheet.write(i+1,4, seller_list[i])
    for i in range(len(ratings_list)):        
            data_worksheet.write(i+1,5, ratings_list[i])
    for i in range(len(reviews_list)):        
            data_worksheet.write(i+1,6, reviews_list[i])  
    data_file.close() 