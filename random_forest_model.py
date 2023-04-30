from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from dtreeviz.trees import dtreeviz # will be used for tree visualization
#from sklearn.tree import export_graphviz
from sklearn import tree
import pandas as pd
import numpy  as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import random
import statsmodels.api as sm
from tabulate import tabulate




###################################################################################################################
for stock in ["2388", "2498"]:
    df = pd.read_excel(f"C:\\Users\\user.Y220026097\\Desktop\\UBS\\{stock}_preprocessed.xlsx", engine="openpyxl")
    df = df.dropna()

    #df = df.drop(df[ df["marketShare"] < 0.05 ].index)
    #print(tabulate(df[df["extreme"]!=0]))

    df = df[["diffPcnt1", "diffPcnt2", "diffPcnt3", "diff1", "diff2", "diff3",
             "totalExcess1", "totalExcess2", "totalExcess3", "daysSoFar", "change1",
             "buy1", "buy2", "buy3", "sell1", "sell2", "sell3", "increaseWithin", "decreaseWithin",
             "extreme", "excessBuy", "increment"]]
    df = df.dropna()
    #X  = df[["diffPcnt1", "diffPcnt2", "diffPcnt3", "diff1", "diff2", "diff3",
    #         "totalExcess1", "totalExcess2", "totalExcess3", "daysSoFar", "change1",
    #        "buy1", "buy2", "buy3", "sell1", "sell2", "sell3", "increaseWithin", "decreaseWithin"]]
    #X = df[["diffPcnt1", "diff1", "totalExcess1", "daysSoFar", "change1",
    #        "buy1", "sell1", "increaseWithin", "decreaseWithin"]]
    X = df[["buy1", "buy2", "buy3", "sell1", "sell2", "sell3", "diffPcnt1", "diffPcnt2", "diffPcnt3"]]

    y  = df[["extreme", "increment"]]
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=0)
    rf = RandomForestClassifier(max_depth=5, random_state=0, criterion="gini", class_weight={0:1,1:10,2:10})
    rf.fit(X_train, y_train["extreme"])
    predictions = rf.predict(X_test)
    wrong       = 0
    for pred, true_val, val in zip(predictions, y_test["extreme"].tolist(), y_test["increment"].tolist()):
        if(pred != true_val):
            wrong += 1
            print(f"true value is {true_val}, but the pred is {pred}. The increment is {val}.")
    print(rf.estimators_[0].tree_.max_depth)
    print(f"the accuracy is { 100 * ( 1 - (wrong/len(predictions) ) ) }%", end="\n\n\n")
    viz = dtreeviz(rf.estimators_[0], X_train, y_train["extreme"], feature_names=X.columns, fontname='Sans')
    viz.save(f"C:\\Users\\user.Y220026097\\Desktop\\UBS\\{stock}_extreme_decision_tree.svg")

    y = df[["excessBuy", "increment"]]
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=0)
    rf = RandomForestClassifier(min_samples_leaf=3, random_state=0, criterion="gini")
    rf.fit(X_train, y_train["excessBuy"])
    predictions = rf.predict(X_test)
    wrong = 0
    for pred, true_val, val in zip(predictions, y_test["excessBuy"].tolist(), y_test["increment"].tolist()):
        if (pred != true_val):
            wrong += 1
            print(f"true value is {true_val}, but the pred is {pred}. The increment is {val}.")
    print(rf.estimators_[0].tree_.max_depth)
    print(f"the accuracy is {100 * (1 - (wrong / len(predictions)))}%", end="\n\n\n")
    viz = dtreeviz(rf.estimators_[0], X_train, y_train["excessBuy"], feature_names=X.columns, fontname='Sans')
    viz.save(f"C:\\Users\\user.Y220026097\\Desktop\\UBS\\{stock}_excessBuy_decision_tree.svg")

