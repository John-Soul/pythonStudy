import pandas as pd
import requests
import json
import codecs

df = pd.read_excel('C:/Users/chian/Desktop/address2latitude/qianniao.xls')
data = df['address'].values
son = data[:1]
errorCount = 0

def address2latitude(item, index):
    if item != 0:
        print(item)
        url = "http://api.map.baidu.com/geocoding/v3/"
        params = {
            'ak': 'Bz0GWjvzAe95YRb9y5pW3UdYwngbTiUN',
            'output': 'json',
            # 'callback': 'showLocation',
            'address': item
        }
        r = requests.get(url, params=params)
        obj = json.load(r.text)
        if obj.status == 0:
            print(obj.result.loaction)
    else:
        str = '暂无数据'

for index in range(len(son)):
    address2latitude(son[index], index)
