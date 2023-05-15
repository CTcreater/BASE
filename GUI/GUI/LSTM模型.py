# -*- coding: utf-8 -*-
"""
Created on Sat May  7 11:03:28 2022

@author: LV
"""

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import tensorflow as tf
from sklearn.preprocessing import StandardScaler
from tensorflow.keras import layers,Input,optimizers
#导入数据
pd.set_option('display.unicode.east_asian_width',True)
data=pd.read_csv(r'C:\Users\Administrator\Desktop\lstm_data\zgpa_train.csv')
df=pd.DataFrame(data,columns=['date','close'])
print(df.head)
#单维单步（用前（1,2,3，......）步预测后1步）
#创建数据集
dataset=df["close"].values
print(type(dataset))
# lookback=1
# print(len(dataset))
def data_set(dataset,lookback):#创建时间序列数据样本
  dataX,dataY=[],[]
  for i in range(len(dataset)-lookback-1):
        a=dataset[i:(i+lookback)]
        dataX.append(a)
        dataY.append(dataset[i+lookback])
  return np.array(dataX),np.array(dataY)


plt.figure(figsize=(12, 8))
x=df["date"]
y=df["close"]
plt.plot(x,y)
plt.show()

st = StandardScaler()
dataset_st = st.fit_transform(dataset.reshape(-1, 1))
print("标准化", dataset_st)
#数据标准化


#划分训练集和测试集
train_size=int(len(dataset_st)*0.7)
test_size=len(dataset_st)-train_size
train,test=dataset_st[0:train_size],dataset_st[train_size:len(dataset_st)]
print(len(train))
print(len(test))

#根据划分的训练集测试集生成需要的时间序列样本数据
lookback=1
trainX,trainY=data_set(train,lookback)
testX,testY=data_set(test,lookback)
print('trianX:',trainX.shape,trainY.shape)
print(trainX)

#构建lstm模型
input_shape=Input(shape=(trainX.shape[1],trainX.shape[2]))
lstm1=layers.LSTM(32,return_sequences=1)(input_shape)
print("lstm1:",lstm1.shape)
lstm2=layers.LSTM(64,return_sequences=0)(lstm1)
print("lstm2:",lstm2.shape)
dense1=layers.Dense(64,activation="relu")(lstm2)
print("dense:",dense1.shape)
dropout=layers.Dropout(rate=0.2)(dense1)
print("dropout:",dropout.shape)
ouput_shape=layers.Dense(1,activation="relu")(dropout)
lstm_model=tf.keras.Model(input_shape,ouput_shape)
lstm_model.compile(loss="mean_squared_error",optimizer="Adam",metrics=["mse"])
history=lstm_model.fit(trainX,trainY,batch_size=16,epochs=10,validation_split=0.1,verbose=1)

lstm_model.summary()

#预测测试集
predict_trainY=lstm_model.predict(trainX)
predict_testY=lstm_model.predict(testX)
#反标准化
#trainY=st.inverse_transform(predict_trainY)
testY_real=st.inverse_transform(testY)
testY_predict=st.inverse_transform(predict_testY)
print("Y:",testY_predict,testY_predict.shape)
print("Y222:",testY_real,testY_real.shape)
plt.figure(figsize=(12,8))
plt.plot(testY_predict,"b",label="预测值")
plt.plot(testY_real,"r",label="真实值")
plt.legend()
plt.show()
