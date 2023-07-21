
# How to create connection and scrape NSE DATA


def establish_session(baseurl,url,headers):
    request=session.get(baseurl,headers=headers)
    cookies=dict(request.cookies)
    response=session.get(url,headers=headers,cookies=cookies).json()
    return response



def dataframe(rawop):
    data=[]
    for i in range(0,len(rawop)):
        calloi=callcoi=cltp=putoi=putcoi=pltp=0
        stp=rawop['strikePrice'][i]
        if(rawop['CE'][i]==0):
            calloi=callcoi=0
        else:
            calloi=rawop['CE'][i]['openInterest']
            callcoi=rawop['CE'][i]['changeinOpenInterest']
            cltp=rawop['CE'][i]['lastPrice']
            calliv=rawop['CE'][i]['impliedVolatility']
            strikeprice=rawop['CE'][i]['strikePrice']
            cltpchng=rawop['CE'][i]['pChange']
            clpricechng=rawop['CE'][i]['change']
        if(rawop['PE'][i])==0:
            putoi=putcoi=0
        else:
            putoi=rawop['PE'][i]['openInterest']
            putcoi=rawop['PE'][i]['changeinOpenInterest']
            pltp=rawop['PE'][i]['lastPrice']
            putiv=rawop['PE'][i]['impliedVolatility']
            putltpchng=rawop['PE'][i]['pChange']
            putpricechng=rawop['PE'][i]['change']            
        opdata={
            'Oi_chng':callcoi,'Oi':calloi,'Iv':calliv, 'Pchange':cltpchng,'PriceChng':clpricechng,'Ltp':cltp,'Strike Price':strikeprice,
            'P_Ltp':pltp,'P_PriceChng':putpricechng,'P_Pchange':putltpchng,'P_Iv':putiv,'P_Oi':putoi,'P_Oi Chng':putcoi
        }
        data.append(opdata)
    optionchain=pd.DataFrame(data)
    return optionchain


def stock_dataframe_stock(rawstock):
    data=[]
    for i in range(0,len(rawstock)):
        symbol=rawstock[i]['symbol']
        opens=rawstock[i]['open']
        dayHigh=rawstock[i]['dayHigh']
        dayLow=rawstock[i]['dayLow']
        lastprice=rawstock[i]['lastPrice']
        previousClose=rawstock[i]['previousClose']
        pChange=rawstock[i]['pChange']
        yearHigh=rawstock[i]['yearHigh']
        yearLow=rawstock[i]['yearLow']
        stock_dict={
            'symbol':symbol,'open':opens,'dayHigh':dayHigh,'dayLow':dayLow,'lastprice':lastprice,'previousClose':lastprice,
            'previousClose':previousClose,'pChange':pChange,'yearHigh':yearHigh,'yearLow':yearLow
        }
        data.append(stock_dict)
    stock_dataframe=pd.DataFrame(data)
    return stock_dataframe


def stock_data():
    response=establish_session(baseurl,url,headers)
    rawdata=pd.DataFrame(response)
    rawop=pd.DataFrame(rawdata['filtered']['data']).fillna(0)
    optionchain=dataframe(rawop)

    session = requests.Session()
    request_2=establish_session(baseurl,fo_url,headers)
    rawstock=request_2['data']
    stock_dataframe2=stock_dataframe_stock(rawstock)
    stock_dataframe2.set_index("symbol", inplace=True)
    future_stock.range("A2").value=stock_dataframe2
    

def option_chain():
    stock_name=option_chain_sheet.range("A2").value
    if stock_name!='NIFTY' and stock_name!='BANKNIFTY':
        build_url=f"https://www.nseindia.com/api/option-chain-equities?symbol={stock_name}"
    else:
        build_url=f"https://www.nseindia.com/api/option-chain-indices?symbol={stock_name}"
    response=establish_session(baseurl,build_url,headers)
    rawdata=pd.DataFrame(response)
    call_total_oi=rawdata['filtered']['CE']['totOI']
    put_total_oi=rawdata['filtered']['PE']['totOI']
    call_total_vol=rawdata['filtered']['CE']['totVol']
    put_total_vol=rawdata['filtered']['PE']['totVol']
    spot_price=rawdata['records']['underlyingValue']
    expiry_date=rawdata['records']['expiryDates'][0]
    option_chain_sheet.range("D2").value=call_total_oi
    option_chain_sheet.range("E2").value=put_total_oi
    option_chain_sheet.range("F2").value=call_total_vol
    option_chain_sheet.range("G2").value=put_total_vol
    option_chain_sheet.range("C2").value=spot_price
    option_chain_sheet.range("J2").value=expiry_date
    rawop=pd.DataFrame(rawdata['filtered']['data']).fillna(0)
    optionchain=dataframe(rawop)
    optionchain.set_index("Oi_chng",inplace=True)
    option_chain_sheet.range("A6").value=optionchain


