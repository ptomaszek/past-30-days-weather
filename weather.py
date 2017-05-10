import requests
import argparse
import time
import xlsxwriter
import datetime


api_key = 'abc123'


# copied and amended from https://github.com/apixu/apixu-python/tree/master/apixu
class ApixuException(Exception):
    def __init__(self, message, code):
        self.message = message
        self.code = code
        message = 'Error code %s: "%s"' % (code, message)
        super(ApixuException, self).__init__(message)


class ApixuClient:
    def __init__(self, api_key=None, host_url='http://api.apixu.com'):
        self.api_key = api_key
        self.host_url = host_url.rstrip('/')

    def _get(self, url, args=None):
        new_args = {}
        if self.api_key:
            new_args['key'] = self.api_key
        new_args.update(args or {})
        response = requests.get(url, params=new_args)
        #print(response.url)
        json_res = response.json()
        if 'error' in json_res:
            err_msg = json_res['error'].get('message')
            err_code = json_res['error'].get('code')
            raise ApixuException(message=err_msg, code=err_code)

        return json_res

    def getHistoricalWeather(self, q, dt, hour):
        url = '%s/v1/history.json' % self.host_url
        args = {}
        args['lang'] = 'pl'
        
        args['q'] = q
        args['dt'] = dt
        args['hour'] = hour
        
        return self._get(url, args)
    
    
    
def toDateStr(date):
    return date.strftime("%Y-%m-%d")
    
def toDate(dateStr):
    return datetime.datetime.strptime(dateStr, "%Y-%m-%d").date()
    
today = datetime.date.today()

query = input('Miejscowosc [Żyrzyn] : ') or 'Żyrzyn'
countBack = int(input('Z ilu ostatnich dni [29] : ') or 29)
print()

fromDate = today - datetime.timedelta(days=countBack)


client = ApixuClient(api_key)

def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + datetime.timedelta(n)
		
def toMps(kph):
    return "{0:.2f}".format(kph / 3.6)
    

workbook = xlsxwriter.Workbook('{}_{}_{}.xlsx'.format(query, toDateStr(fromDate), toDateStr(today)));
worksheet = workbook.add_worksheet()

format = workbook.add_format()
format.set_bold()

worksheet.write(0, 0, query, format)


startCol = 2
worksheet.write(0, startCol, 'Data', format)
worksheet.write(0, startCol+1, 'Min. temp. [°C]', format)
worksheet.write(0, startCol+2, 'Maks. temp. [°C]', format)
worksheet.write(0, startCol+3, 'Śr. temp. [°C]', format)
worksheet.write(0, startCol+4, 'Opady [mm]', format)
worksheet.write(0, startCol+5, 'Wiatr [m/s]', format)
worksheet.write(0, startCol+6, 'Pogoda', format)
worksheet.write(0, startCol+7, '')


worksheet.set_column(0, 0, 15)
worksheet.set_column(2, 2, 15)
worksheet.set_column(3, 3, 15)
worksheet.set_column(4, 4, 15)
worksheet.set_column(5, 5, 15)
worksheet.set_column(6, 6, 15)
worksheet.set_column(7, 7, 15)
worksheet.set_column(8, 8, 32)


row = 1

from io import BytesIO
from urllib.request import urlopen

cell_format = workbook.add_format({})


maxWidths = []

for date in daterange(fromDate, today):
    dateStr = toDateStr(date)
    print ('Pobieram dane pogodowe z dnia {}...'.format(dateStr))
    col = startCol
    historical = client.getHistoricalWeather(q=query, dt=dateStr, hour=13) 
    
    data = [
        str(historical['forecast']['forecastday'][0]['date']),
        str(historical['forecast']['forecastday'][0]['day']['mintemp_c']),
        str(historical['forecast']['forecastday'][0]['day']['maxtemp_c']),
        str(historical['forecast']['forecastday'][0]['day']['avgtemp_c']),
        str(historical['forecast']['forecastday'][0]['day']['totalprecip_mm']),
        str(toMps(historical['forecast']['forecastday'][0]['day']['maxwind_kph'])),
        str(historical['forecast']['forecastday'][0]['day']['condition']['text']),
        str(historical['forecast']['forecastday'][0]['day']['condition']['icon'])
    ]
    
    for value in data:
        if '.png' in value:
            image_data = BytesIO(urlopen('http:' + value).read())
            worksheet.insert_image(row, col, 'http:' + value, {'image_data': image_data, 'y_offset': -8, 'x_scale': 0.7, 'y_scale': 0.7})
        else:
            worksheet.write(row, col, value)
        
        col=col+1

    worksheet.set_row(row, 20)
    row=row+1
     
print ('\nOK!')

workbook.close()    


