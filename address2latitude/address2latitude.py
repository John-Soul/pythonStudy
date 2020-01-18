import pandas as pd
import requests
import json
from pandas import DataFrame
from pandas import ExcelWriter

df = pd.read_excel('C:/Users/chian/Desktop/pythonStudy/qianniao.xls')
companyData = df['company'].values
addressData = df['address'].values
weightData = df['weight'].values
longitudeData = []
latitudeData = []

def address2latitude(item, index):
    if item != 0:
        url = "http://api.map.baidu.com/geocoding/v3/"
        params = {
            'ak': 'Bz0GWjvzAe95YRb9y5pW3UdYwngbTiUN',
            'output': 'json',
            'address': item
        }
        r = requests.get(url, params=params)
        obj = json.loads(r.text)
        if obj.get('status') == 0:
            location = obj.get('result').get('location')
            latitudeData.append(location.get('lat'))
            longitudeData.append(location.get('lng'))
            
    else:
        str = '请提供精确的地理位置'
        latitudeData.append(str)
        longitudeData.append(str)

# def toJSON():


def toExcel():
    df['longitude'] = pd.DataFrame(longitudeData)
    df['latitude'] = pd.DataFrame(latitudeData)
    DataFrame(df).to_excel('C:/Users/chian/Desktop/pythonStudy/result.xls', sheet_name='result', index=True, header=True)

for index in range(len(addressData)):
    address2latitude(addressData[index], index)
    if index + 1 == len(addressData):
        toExcel()
