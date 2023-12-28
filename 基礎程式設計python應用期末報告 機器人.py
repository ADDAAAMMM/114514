import time  ## 匯入 time 模組，用於計算程式執行時間
import requests  ## 匯入 requests 模組，用於發送 HTTP 請求
from pandas import DataFrame  ## 匯入 DataFrame 類別，用於建立資料表格
import json  ## 匯入 json 模組，用於處理 JSON 資料

def main():  ## 定義主函式
    print('\n===============程式開始======================\n')  ## 輸出程式開始的訊息
    keyword=input('請輸入欲查詢之商品的關鍵字... ：')  ## 輸入欲查詢的商品關鍵字

    url = 'https://ecshweb.pchome.com.tw/search/v3.3/all/results'  ## 指定要查詢商品的網址

    data = Get_PageContent(url, keyword, 1)  ## 使用 Get_PageContent 函式獲取第 1 頁的商品資料
    total_page_num = int(int(data['totalRows'])/20)+1  ## 計算總頁數
    print('\n查詢結果約有 {} 頁，共{}筆資料。'.format(total_page_num, int(data['totalRows'])))  ## 輸出查詢結果的訊息
    page_want_to_crawl = input('一頁有20筆，請問你要爬取多少頁? ')  ## 輸入欲爬取的頁數
    if page_want_to_crawl == '' or not page_want_to_crawl.isdigit() or int(page_want_to_crawl) <= 0:  ## 若輸入的頁數不合法，則顯示錯誤訊息
        print('\n頁數輸入錯誤，離開程式')
        print('\n===============程式結束======================\n')
    else:  ## 否則執行以下程式碼
        page_want_to_crawl = min(int(page_want_to_crawl), int(total_page_num))  ## 取輸入的頁數和總頁數中的較小值
        print('\n計算中，請稍候。。。。。')
        start = time.time()  ## 記錄程式開始執行時間
        products = Parse_Get_MetaData(url, keyword, page_want_to_crawl)  ## 使用 Parse_Get_MetaData 函式獲取指定頁數的商品資料
        print('\n已取得所需商品，執行時間共 {} 秒。'.format(time.time()-start))  ## 輸出程式執行時間
        Save2Excel(products)  ## 使用 Save2Excel 函式將商品資料存入 Excel 檔案
        print('\n====資料已順利取得，並已存入pchome24.xlsx中====\n')  ## 輸出資料存檔成功的訊息

def Get_PageContent(url, keyword, i):  ## 定義 Get_PageContent 函式，用於發送 HTTP 請求獲取指定頁數的商品資料
    my_params = {  ## 建立參數字典
        'q': keyword,  ## 指定關鍵字
        'page': i,  ## 指定頁數
        'sort': 'sale/dc'  ## 指定排序方式
        }
    res = requests.get(url, params = my_params)  ## 發送 GET 請求
    content = json.loads(res.text)  ## 將回傳的 JSON 資料解析為 Python 物件
    print(content)  ## 輸出回傳的資料
    return content  ## 回傳資料

def Parse_Get_MetaData(url, keyword, page):  ## 定義 Parse_Get_MetaData 函式，用於解析指定頁數的商品資料
    products_list = list()  ## 建立空的商品列表
    product_no = 0  ## 初始商品編號為 0

    for i in range(1,page+1):  ## 迴圈從 1 到指定頁數
        data = Get_PageContent(url, keyword, i)  ## 使用 Get_PageContent 函式獲取商品資料
        if 'prods' in data:  ## 若資料中存在 'prods' 鍵
            products = data['prods']  ## 獲取商品列表

            for product in products:  ## 迴圈遍歷每個商品
                product_no +=1  ## 商品編號加 1
                products_list.append({  ## 將商品資訊以字典形式加入商品列表
                                '編號': product_no,
                                '品名': product['name'],
                                '商品連結': 'https://24h.pchome.com.tw/prod/'+ product['Id'],
                                '價格': product['price']
                                })        
        else:  ## 若資料中不存在 'prods' 鍵，則跳出迴圈
            break  
    print(products_list)  ## 輸出商品列表
    return products_list  ## 回傳商品列表

def Save2Excel(products):  ## 定義 Save2Excel 函式，將商品資料存入 Excel 檔案
    product_no = [entry['編號'] for entry in products]  ## 獲取商品編號列表
    product = [entry['品名'] for entry in products]  ## 獲取商品名稱列表
    product_link = [entry['商品連結'] for entry in products]  ## 獲取商品連結列表
    price = [entry['價格'] for entry in products]  ## 獲取商品價格列表

    df = DataFrame({  ## 建立 DataFrame 物件，包含商品編號、商品名稱、商品連結和商品價格四個欄位
        '編號':product_no,
        '品名':product,
        '商品連結':product_link,
        '價格':price
        })
    df.to_excel('pchome24.xlsx', sheet_name='sheet1', columns=['編號','品名', '商品連結', '價格'])  ## 將資料存入 Excel 檔案，指定欄位順序

if __name__ == '__main__':  ## 若為主程式執行，則執行 main 函式
    main()