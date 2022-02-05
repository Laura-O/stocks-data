import json
import certifi
import os
from urllib.request import urlopen

API_KEY = os.environ.get("FMP_KEY")

profile_url = 'https://financialmodelingprep.com/api/v3/profile/%s?apikey=' +  API_KEY
rating_url = 'https://financialmodelingprep.com/api/v3/rating/%s?apikey=' + API_KEY
income_url = 'https://financialmodelingprep.com/api/v3/income-statement/%s?period=quarter&limit=400&apikey=' + API_KEY

def get_jsonparsed_data(url):
    response = urlopen(url, cafile=certifi.where())
    data = response.read().decode("utf-8")
    return json.loads(data)

class ProfileFMP(object):
    def __init__(self, symbol):
        self.symbol = symbol
        self.set_profile(symbol)
        self.set_rating(symbol)
        self.set_income(symbol)

    def set_profile(self, symbol):        
        data = get_jsonparsed_data(profile_url % symbol)
        self.profile = data[0]

    def set_rating(self, symbol):        
        data = get_jsonparsed_data(rating_url % symbol)
        self.rating = data[0]

    def set_income(self, symbol):        
        data = get_jsonparsed_data(income_url % symbol)
        self.income = data