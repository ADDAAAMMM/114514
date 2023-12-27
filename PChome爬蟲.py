import time
import requests
from pandas import DataFrame
import json

##############################
# 主程式-- 取得PChome24的商品  #
##############################

def main():
    print('\n===============程式開始======================\n')
    keyword=input('請輸入欲查詢之商品的關鍵字... ：')

    url = 'https://ecshweb.pchome.com.tw/search/v3.3/all/results'

    data = Get_PageContent(url, keyword, 1)
    total_page_num = int(int(data['totalRows'])/20)+1

    print('\n查詢結果約有 {} 頁，共{}筆資料。'.format(total_page_num, int(data['totalRows'])))

    page_want_to_crawl = input('一頁有20筆，請問你要爬取多少頁? ')
    if page_want_to_crawl == '' or not page_want_to_crawl.isdigit() or int(page_want_to_crawl) <= 0:
        print('\n頁數輸入錯誤，離開程式')
        print('\n===============程式結束======================\n')
    else:
        page_want_to_crawl = min(int(page_want_to_crawl), int(total_page_num))

        print('\n計算中，請稍候。。。。。')
        start = time.time()
        products = Parse_Get_MetaData(url, keyword, page_want_to_crawl)
        print('\n已取得所需商品，執行時間共 {} 秒。'.format(time.time()-start))
        Save2Excel(products)
        print('\n====資料已順利取得，並已存入pchome24.xlsx中====\n')


##########################################
# 提出請求，取得頁面資料(JSON格式)         #
##########################################
def Get_PageContent(url, keyword, i):
    my_params = {
        'q': keyword,
        'page': i,
        'sort': 'sale/dc'
        }
    res = requests.get(url, params = my_params)
    content = json.loads(res.text)
    print(content)
    return content


############################################################
# 取得各頁面的商品，然後統整成包含所有商品的串列。             #
############################################################
def Parse_Get_MetaData(url, keyword, page):
    products_list = list()
    product_no = 0

    #依頁碼順序取資料，各頁的商品包在'prods'中
    for i in range(1,page+1): 
        data = Get_PageContent(url, keyword, i)
        if 'prods' in data:
            products = data['prods']

            #取出各頁中的商品
            for product in products:
                product_no +=1
                products_list.append({
                                '編號': product_no,
                                '品名': product['name'],
                                '商品連結': 'https://24h.pchome.com.tw/prod/'+ product['Id'],
                                '價格': product['price']
                                })        
        else:
            break  
    print(products_list)  
    return products_list


#####################################
# 將商品資料存入pchome24.xlsx 中      #
#####################################
def Save2Excel(products):
    product_no = [entry['編號'] for entry in products]
    product = [entry['品名'] for entry in products]
    product_link = [entry['商品連結'] for entry in products]
    price = [entry['價格'] for entry in products]

    df = DataFrame({
        '編號':product_no,
        '品名':product,
        '商品連結':product_link,
        '價格':price
        })
    df.to_excel('pchome24.xlsx', sheet_name='sheet1', columns=['編號', '品名', '商品連結', '價格'])   


#####################################
# 設定程式可被 import或直譯           #
#####################################        
if __name__ == '__main__':
    main()