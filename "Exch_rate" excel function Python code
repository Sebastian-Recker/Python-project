# This imports the xlwings library
import xlwings as xw
# This imports the UrlLib library, used to fetch the currency exchange rate through the internet
import urllib

# This defines a new Excel function
@xw.func
# This defines the name of the Excel function and the two variables wich will Have to be entered by the Excel user
def exch_rate(from_c,to_c):
    # Here we fetch the conversion rate between the two user specified currencies
    url = urllib.request.urlopen("https://www.exchangerate-api.com/%s/%s?k=9644d5faaa0392612a451294"%(from_c,to_c))
    # Here we store the result in the "result" variable
    result = url.read()
    # Here we send the exchange rate to Excel, [2:] removes the two first characters and [:-1] removes the last one, this is done for formating purposes
    return (str(result)[2:])[:-1]
