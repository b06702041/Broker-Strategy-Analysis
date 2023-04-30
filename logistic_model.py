from sklearn.linear_model import LogisticRegressionCV
import pandas as pd
import numpy  as np
import matplotlib.pyplot as plt
import random
import statsmodels.api as sm
import statsmodels.formula.api as smf
from tabulate import tabulate

for stock in ["2388", "2498"]:
    df = pd.read_excel(f"C:\\Users\\user.Y220026097\\Desktop\\UBS\\{stock}_preprocessed.xlsx", engine="openpyxl")
    df = df.dropna()
    #print(tabulate(df))

    X_dataset = df[ ["opening", "highest", "lowest", "closing", "priceDiff", "diffPercent", "volume",
                     "marketShare", "days", "acmlPercent", "totalExcess1", "totalExcess2", "totalExcess3",
                     "positive3", "negative3", "buy1", "buy2", "buy3", "sell1", "sell2", "sell3"] ]
    y_dataset = df["increment"]

    num_var = [ "increaseWithin", "decreaseWithin", "daysSoFar", "change1",
                "totalExcess1", "totalExcess2", "totalExcess3",
                "diffPcnt1", "diffPcnt2", "diffPcnt3", "diff1", "diff2", "diff3",
                "buy1", "buy2", "buy3", "sell1", "sell2", "sell3"]

    #dummy   = ["positive3", "negative3"]



    """OLS Regression"""
    #ols_model  = smf.ols(formula=f"increment  ~ {'+'.join(num_var)} + C(positive3) + C(negative3)", data=df)
    ols_model = smf.ols(formula=f"increment ~ {'+'.join(num_var)} + C(positive2) + C(negative2)", data=df)
    result    = ols_model.fit()
    print(result.summary2())

    ols_model = smf.ols(formula=f"incrementDiff ~ {'+'.join(num_var)} + C(positive2) + C(negative2)", data=df)
    result = ols_model.fit()
    print(result.summary2())

    """Logistic Regression"""
    y_dataset   = df["excessBuy"]
    df0_X       = sm.add_constant(df[num_var])
    logit_model = sm.Logit(y_dataset, df0_X)
    result      = logit_model.fit()
    print(result.summary2())
