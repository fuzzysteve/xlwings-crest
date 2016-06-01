@xl.func    
def crestHighestBuyPrice(typeid,regionid):
    response=requests.get('https://crest-tq.eveonline.com/market/{}/orders/buy/?type=https://crest-tq.eveonline.com/types/{}/'.format(int(regionid),int(typeid)))
    data=response.json()
    highestBuyPrice=0
    for order in data['items']:
        if order['price']>highestBuyPrice:
            highestBuyPrice=order['price']
    return highestBuyPrice
