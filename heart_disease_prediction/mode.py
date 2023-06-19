import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score



heart_data=pd.read_csv('D:\college\sem_5\heart_disease\heart.csv')

x= heart_data.drop(columns='target',axis=1)
y= heart_data['target']

x_train,x_test,y_train,y_test = train_test_split(x,y, test_size=0.2, stratify=y, random_state=2)
model=LogisticRegression()
model.fit(x_train,y_train)

input_data=(34,0,1,118,210,0,1,192,0,0.7,2,0,2)

#change input data to numpy array
input_data_array=np.asarray(input_data)

#reshape the array as we are predicting for only on instance
input_data_reshape=input_data_array.reshape(1,-1)

prediction=model.predict(input_data_reshape)
if(prediction[0]==0):
    print("person have healthy heart")
    
else:
    print("person have unhealthy heart")