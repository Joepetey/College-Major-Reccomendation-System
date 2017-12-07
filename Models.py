# -*- coding: utf-8 -*-
"""
Created on Sun Dec  3 22:11:20 2017

@author: goldw
"""
import numpy as np
from sklearn import svm
import pandas as pd
from sklearn.naive_bayes import BernoulliNB
from sklearn.model_selection import train_test_split
from sklearn.utils import shuffle
from sklearn.metrics import accuracy_score

data = pd.read_csv('C:\\Users\\goldw\\Documents\\PrototypeMajorPathData.csv',delimiter = ',')
data = shuffle(data)

labels = data.ix[:,0] # shuffled labels
features = data.drop('Commercial Art & Design1', 1) # binary feature dataframe
train, test, train_labels, test_labels = train_test_split(features, labels, test_size = 0.30, random_state = 42)

bnb = BernoulliNB()
bnb.fit(train, train_labels)
svm = svm.SVC(decision_function_shape='ovo', degree = 3)
svm.fit(train, train_labels)

bnbpred = bnb.predict(test)
svmpred = svm.predict(test)

print(accuracy_score(test_labels, bnbpred)*100)
print(accuracy_score(test_labels, svmpred)*100)
