import pandas as pd
import json

excelData = pd.read_excel('C:/Users/chian/Desktop/pythonStudy/result.xls')
companyData = excelData['company'].values
addressData = excelData['address'].values
weightData = excelData['weight'].values
longitudeData = excelData['longitude'].values
latitudeData = excelData['latitude'].values

data = []
geo = {}

def e2l(item, index):
    if addressData[index] != 0:
        if isinstance(item, float):
            print(item)
            data.append({
                'name': companyData[index],
                'value': weightData[index]
            })
            geo[companyData[index]] = [longitudeData[index], latitudeData[index]]


def toJSON():
    with open('C:/Users/chian/Desktop/pythonStudy/data.json', 'w') as DataResult:
        json.dump(data, DataResult, ensure_ascii=False)

    with open('C:/Users/chian/Desktop/pythonStudy/geo.json', 'w') as GeoResult:
        json.dump(geo, GeoResult, ensure_ascii=False)

for index in range(len(longitudeData)):
    e2l(longitudeData[index], index)
    if index + 1 == len(longitudeData):
        toJSON()