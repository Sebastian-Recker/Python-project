# This imports the xlwings library
import xlwings as xw
# This imports the UrlLib library, used to fetch the currency exchange rate through the internet
import urllib
#This import the JSON library, used to handle the list-like item we get from the fixer.io exchange-rate website
import json

#This defines a new Excel function  
@xw.func
# This defines the name of the Excel function and the three variables wich will have to be entered by the Excel user
def EXCH(date,from_c,to_c):
    # Here we fetch the conversion rate between the two user specified currencies
    url = urllib.request.urlopen("http://api.fixer.io/%s?symbols=%s,%s"%(date,from_c,to_c))
    # Here we store the result in the "url_res" variable and change its formating from bytes to utf8, this is necessary for it to be imported as a JSON object
    url_res = url.read().decode('utf-8')
    #Here we load the URL result in a JSON format
    json_dump = json.loads(url_res)
    #We now extract the "rates" part of the JSON list
    json_rates = json_dump["rates"]
    #We extract the target currency from the "rates" list
    json_to_cur = json_rates[to_c]
    #Finally, we return the target currency rate at the user specified date to excel
    return json_to_cur
