import numpy as np
import sys
import pandas as pd
def ahp_weights():

    data=pd.read_excel("ahp_in.xlsx",header=None)
    print("AHP matrix is:\n",data)
    data.loc[len(data),:] = [np.NAN for i in range(len(data.columns))]
    #print(data)
    col_len=len(data.columns)
    row_len=len(data)
    data.iloc[row_len-1,0]='sum'
    for j in range(1,col_len):
        sum=0
        for i in range(1,row_len-1):
            sum+=data.iloc[i,j]
        data.iloc[row_len-1,j]=sum
    #print(data)
    for j in range(1,col_len):
        for i in range(1,row_len-1):
            data.iloc[i,j]/=data.iloc[row_len-1,j]
    print("All computation for weight matrix of AHP:\n",data)
    #wt = pd.DataFrame({'weights': [np.NaN for i in range(row_len-1)]})
    #print(wt)
    wt = data.iloc[0:row_len-1,0:2]
    wt.iloc[0,1]='weights'
    #print(wt)

    for i in range(1,row_len-1):
        sum=0
        for j in range(1,col_len):
            sum+=data.iloc[i,j]
        wt.iloc[i,1]=sum/(col_len-1)
    print("The weight matrix is: ",wt)
    wt.to_excel('ahp_weights.xlsx',index=False,header=False)
if __name__=="__main__":
    ahp_weights()
