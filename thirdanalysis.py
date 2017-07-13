from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
import xml.etree.ElementTree as ET
tree = ET.parse('shop.xml')
root = tree.getroot()


productid = input("Pleasue enter ProductIDType: ")

print("Enter ProductIDType : ",productid)

productdetials = []

for child in root.findall('Product'):
    DistinctiveTitle = child.find('DistinctiveTitle').text

    for id in child.findall('ProductIdentifier'):
        ProductIDType = id.find('ProductIDType').text
        if ProductIDType == '15':
            IDValue = id.find('IDValue').text
            pricedetails = []
            details = {"DistinctiveTitle": DistinctiveTitle,"IDValue":IDValue,"pricedetails":pricedetails}
            productdetials.append(details)

            SupplyDetail = child.find('SupplyDetail')
            for price in SupplyDetail.findall('Price'):
                CurrencyCode = price.find('CurrencyCode').text
                if CurrencyCode == 'GBP':
                    gbp = price.find('PriceAmount').text
                    gbp1 = float(gbp) * 1.13 #GBP to EUR google rate
                    eur = SupplyDetail.findall('Price')[1].find('PriceAmount').text
                    decision = []
                    if gbp1 > float(eur):
                        detailstwo = {'gbp': gbp, "converted": gbp1, "eur": eur, "Cheaper": "EUR"}
                        pricedetails.append(detailstwo)
                    else:
                        detailstwo = {'gbp': gbp, "converted": gbp1, "eur": eur, "Cheaper": "GBP"}
                        pricedetails.append(detailstwo)

# print(productdetials)

for i in productdetials:
    if productid == i['IDValue']:
        # print("Product Id:: ", i['IDValue'])
        print("GBP price:: ",i['pricedetails'][0]['gbp'])
        print("GBP price converted to EUR:: ",i['pricedetails'][0]['converted'])
        print("EUR price:: ",i['pricedetails'][0]['eur'])
        print("Cheaper price:: ",i['pricedetails'][0]['Cheaper'])
        exit()
    else:
        print("Does Not Match Product Id")
        exit()

