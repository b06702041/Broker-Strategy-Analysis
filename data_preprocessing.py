# -*- coding: utf-8 -*-
import os, shutil
import pandas as pd
import numpy  as np
from tqdm import tqdm
from WCFAdox import PCAX
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.font_manager import FontProperties
from tabulate import tabulate

def date_transform(text):
    date = text[:4] + '-' + text[4:6] + '-' + text[6:]
    return date

def clear_construct_folder(foldername):
    try:
        shutil.rmtree(foldername)
    except:
        os.makedirs(foldername)
    else:
        os.makedirs(foldername)

def get_UBS_data(foldername, startDate, endDate, stock):
    PX = PCAX("10.216.8.148")
    df = PX.Sil_Data("個股券商進出明細", "D", stock, startDate, endDate, isst="Y")
    df = df[ df["券商名稱"]=="新加坡商瑞銀" ]
    df["日期"] = df["日期"].astype(str).apply(date_transform)
    df = df.iloc[::-1]

    path = f"{foldername}\\UBS_{stock}.xlsx"
    df.to_excel(path, index=False, encoding="utf-8", engine="xlsxwriter")

def get_stock_data(foldername, startDate, endDate, stock):
    PX = PCAX("10.216.8.148")
    df = PX.Sil_Data("日盤中零股交易行情", "D", stock, startDate, endDate, isst="Y")
    df["日期"] = df["日期"].astype(str).apply(date_transform)
    df["日期"] = pd.to_datetime(df["日期"], format="%Y-%m-%d", yearfirst=True)
    df        = df.sort_values(by="日期")
    df["日期"] = df["日期"].dt.strftime("%Y-%m-%d")

    path = f"{foldername}\\stock_{stock}.xlsx"
    with pd.ExcelWriter(path, engine='xlsxwriter', datetime_format='YYYY-MM-DD', date_format='YYYY-MM-DD') as writer:
        df.to_excel(writer, index=False)

def get_TWII_data(foldername, startDate, endDate):
    PX        = PCAX("10.216.8.148")
    df        = PX.Sil_Data("重要國際指數", "D", "#TWII", startDate, endDate, isst="N")
    df["日期"] = df["日期"].astype(str).apply(date_transform)
    df["日期"] = pd.to_datetime(df["日期"], format="%Y-%m-%d", yearfirst=True)
    df        = df.sort_values(by="日期")
    df["日期"] = df["日期"].dt.strftime("%Y-%m-%d")

    df         = df[["日期", "開盤價", "最高價", "最低價", "收盤價", "漲跌", "漲跌幅(%)", "成交量"]]
    df.columns = ["date", "TWII_opening", "TWII_highest", "TWII_lowest", "TWII_closing",
                  "TWII_priceDiff", "TWII_diffPercent", "TWII_volume"]

    path = f"{foldername}\\TWII.xlsx"
    with pd.ExcelWriter(path, engine='xlsxwriter', datetime_format='YYYY-MM-DD', date_format='YYYY-MM-DD') as writer:
        df.to_excel(writer, index=False)

    ####################################################################################
    df        = PX.Sil_Data("期貨交易行情表", "D", "TX", startDate, endDate, isst="N")
    df["日期"] = df["日期"].astype(str).apply(date_transform)
    df["日期"] = pd.to_datetime(df["日期"], format="%Y-%m-%d", yearfirst=True)
    df        = df.sort_values(by="日期")
    df["日期"] = df["日期"].dt.strftime("%Y-%m-%d")

    df         = df[ ["日期", "開盤價", "最高價", "最低價", "收盤價", "漲跌", "漲幅(%)", "成交量"] ]
    df.columns = ["date", "future_opening", "future_highest", "future_lowest", "future_closing",
                  "future_priceDiff", "future_diffPercent", "future_volume"]

    path = f"{foldername}\\TWII_future.xlsx"
    with pd.ExcelWriter(path, engine='xlsxwriter', datetime_format='YYYY-MM-DD', date_format='YYYY-MM-DD') as writer:
        df.to_excel(writer, index=False)
############################################### Key Function ##############################################
def get_raw_data(stock, dataFolder, startDate, endDate):
    foldername = f"{dataFolder}\\raw data"
    clear_construct_folder(foldername)

    get_UBS_data  (foldername, startDate, endDate, stock)
    get_stock_data(foldername, startDate, endDate, stock)
    get_TWII_data (foldername, startDate, endDate)

def data_selection(stock, dataFolder):
    foldername = f"{dataFolder}\\raw data"
    path1      = f"{foldername}\\stock_{stock}.xlsx"
    path2      = f"{foldername}\\UBS_{stock}.xlsx"
    path3      = f"{dataFolder}\\10year_treasury.xlsx"
    df1 = pd.read_excel(path1, engine="openpyxl")
    df2 = pd.read_excel(path2, engine="openpyxl")
    df3 = pd.read_excel(path3, engine="openpyxl")
    df3["date"] = pd.to_datetime(df3["date"], format="%Y-%m-%d", yearfirst=True)
    df3["date"] = df3["date"].dt.strftime("%Y-%m-%d")

    df = pd.merge(df2, df1, on=["日期"], how='outer')
    df = df[ ["日期", "開盤價", "最高價", "最低價", "收盤價", "漲跌", "漲幅(%)", "成交量(股)",
              "張增減", "買張", "賣張"]]
    df.columns = ["date", "opening", "highest", "lowest", "closing", "priceDiff", "diffPercent", "volume",
                  "increment", "buy", "sell"]

    #df = pd.merge(df3, df, on=["date"], how='outer')
    df = df.sort_values(by="date")

    path = f"{foldername}\\{stock}_selection.xlsx"
    df.to_excel(path, index=False, encoding="utf-8", engine="xlsxwriter")

def data_processing(stock, dataFolder):
    foldername = f"{dataFolder}\\raw data"
    path       = f"{foldername}\\{stock}_selection.xlsx"
    df         = pd.read_excel(path, engine="openpyxl")

    df["excessBuy"] = 1
    df["excessBuy"][ df["increment"].astype(float) < 0] = 0

    """1"""
    df["marketShare"] = (df["buy"] + df["sell"]) / df["volume"]

    """2, 3"""
    df["diffPercent"].astype(float)
    df["days"]        = 0
    df["acmlPercent"] = 0
    acml = 0
    for ind in range(1, len(df.index)):
        if (df.iloc[ind]["diffPercent"] == 0):
            df.at[ind, "days"] = df.iloc[ind - 1]["days"]
            df.at[ind, "acmlPercent"] = acml

        elif (df.iloc[ind]["diffPercent"] > 0 and df.iloc[ind - 1]["days"] >= 0):
            acml += df.iloc[ind]["diffPercent"]
            df.at[ind, "days"] = df.iloc[ind - 1]["days"] + 1
            df.at[ind, "acmlPercent"] = acml
        elif (df.iloc[ind]["diffPercent"] > 0 and df.iloc[ind - 1]["days"] < 0):
            acml = df.iloc[ind]["diffPercent"]
            df.at[ind, "days"] = 1
            df.at[ind, "acmlPercent"] = acml

        elif (df.iloc[ind]["diffPercent"] < 0 and df.iloc[ind - 1]["days"] <= 0):
            acml += df.iloc[ind]["diffPercent"]
            df.at[ind, "days"] = df.iloc[ind - 1]["days"] - 1
            df.at[ind, "acmlPercent"] = acml
        elif (df.iloc[ind]["diffPercent"] < 0 and df.iloc[ind - 1]["days"] > 0):
            acml = df.iloc[ind]["diffPercent"]
            df.at[ind, "days"] = -1
            df.at[ind, "acmlPercent"] = acml
    df["daysSoFar"] = df["days"].shift(1)
    df["change"]    = (df["closing"] / df.iloc[0]["closing"]) - 1
    df["change1"]   = df["change"].shift(1)

    """4"""
    df["-1"] = df["increment"].shift(1)
    df["-2"] = df["increment"].shift(2)
    df["-3"] = df["increment"].shift(3)
    df["totalExcess1"] = df["-1"]
    df["totalExcess2"] = df["-1"] + df["-2"]
    df["totalExcess3"] = df["-1"] + df["-2"] + df["-3"]
    df = df.drop(columns=["-1", "-2", "-3"])

    """5"""
    #buy_std  = df["increment"][df["increment"] > 0].std()
    #sell_std = df["increment"][df["increment"] < 0].std()
    buy_threshold  = df["increment"][df["increment"] > 0].quantile(0.8)
    sell_threshold = df["increment"][df["increment"] < 0].quantile(0.2)
    df["extreme"] = 0 # normal
    df["BUY"]     = 0
    df["SELL"]    = 0
    #df["extreme"][ (df["increment"] > 0) & (df["increment"]      > buy_std  * 1) ] = 1 # "BUY!"
    #df["extreme"][ (df["increment"] < 0) & (abs(df["increment"]) > sell_std * 1) ] = 2 # "SELL!"
    df["extreme"][ (df["increment"] > 0) & (df["increment"] > buy_threshold)  ] = 1  # "BUY!"
    df["extreme"][ (df["increment"] < 0) & (df["increment"] < sell_threshold) ] = 2 # "SELL!"
    df["BUY"][  (df["increment"] > 0) & (df["increment"] > buy_threshold)  ] = 1
    df["SELL"][ (df["increment"] < 0) & (df["increment"] > sell_threshold) ] = 1

    """6"""
    df["positive3"] = 0
    df["negative3"] = 0
    df["positive2"] = 0
    df["negative2"] = 0
    df["positive3"][df["days"] >= 3] = 1
    df["negative3"][df["days"] <= -3] = 1
    df["positive2"][ df["days"].shift(1) >=  2] = 1
    df["negative2"][ df["days"].shift(1) <= -2] = 1

    """7"""
    df["buy1"]  = df["buy"].shift(1,fill_value=0)
    df["buy2"]  = df["buy"].shift(2,fill_value=0)
    df["buy3"]  = df["buy"].shift(3,fill_value=0)
    df["sell1"] = df["sell"].shift(1,fill_value=0)
    df["sell2"] = df["sell"].shift(2,fill_value=0)
    df["sell3"] = df["sell"].shift(3,fill_value=0)

    """8"""
    df["increaseWithin"] = (df["highest"] / df["opening"]) - 1
    df["decreaseWithin"] = (df["lowest"]  / df["opening"]) - 1
    #df["TWII_increaseWithin"] = (df["TWII_highest"] / df["TWII_opening"]) - 1
    #df["TWII_decreaseWithin"] = (df["TWII_lowest"] /  df["TWII_opening"]) - 1
    #df["future_increaseWithin"] = (df["future_highest"] / df["future_opening"]) - 1
    #df["future_decreaseWithin"] = (df["future_lowest"]  / df["future_opening"]) - 1

    """9"""
    df["closing1"]  = df["closing"].shift(1)
    df["closing2"]  = df["closing"].shift(2)
    df["closing3"]  = df["closing"].shift(3)
    df["diffPcnt1"] = (df["opening"] / df["closing1"]) - 1
    df["diffPcnt2"] = (df["opening"] / df["closing2"]) - 1
    df["diffPcnt3"] = (df["opening"] / df["closing3"]) - 1
    df["diff"]  = df["opening"] - df["closing"]
    df["diff1"] = df["opening"] - df["closing1"]
    df["diff2"] = df["opening"] - df["closing2"]
    df["diff3"] = df["opening"] - df["closing3"]

    """10"""
    #df["opening_spread"] = df["future_opening"] - df["TWII_opening"]
    df["incrementDiff"] = df["increment"] - df["increment"].shift(1)
    df["increment1"]    = df["increment"].shift(1)
    df["increment2"]    = df["increment"].shift(2)
    df["increment3"]    = df["increment"].shift(3)

    path = f"{dataFolder}\\{stock}_preprocessed.xlsx"
    df.to_excel(path, index=False, encoding="utf-8", engine="xlsxwriter")

###########################################################################################################


if __name__ == '__main__':

    dataFolder = "C:\\Users\\user.Y220026097\\Desktop\\UBS"
    startDate  = "20211027"
    endDate    = "20221031"

    for stock in ["2388", "2498"]:
        get_raw_data   (stock, dataFolder, startDate, endDate)
        data_selection (stock, dataFolder)
        data_processing(stock, dataFolder)