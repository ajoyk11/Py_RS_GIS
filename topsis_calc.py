import numpy as np
import pandas as pd
import math

def topsis_calc():
    data=pd.read_excel("topsis_in.xlsx",header=None)
    print("The Topsis matrix is:\n",data)
    #print(data)
    col_len=len(data.columns)
    row_len=len(data)
    #print(row_len,col_len)
    new_rows = pd.DataFrame([[np.NAN for i in range(col_len)] for j in range(5)])
    data = data.append(new_rows, ignore_index=True)
    #print(data)
    data.iloc[row_len,0]='SQUARE SUM'
    data.iloc[row_len+1,0]='ROOT SQUARE SUM'
    data.iloc[row_len+2,0]='WEIGHT'
    data.iloc[row_len+3,0]='MAX'
    data.iloc[row_len+4,0]='MIN'
    wt=pd.read_excel("ahp_weights.xlsx",header=None)
    #Transferring weight from ahp matrix
    #print(wt)
    for j in range(1,col_len):
        data.iloc[row_len+2,j]=wt.iloc[j,1]
        sq_sum=0
        for i in range(1,row_len-1):
            sq_sum+=data.iloc[i,j]*data.iloc[i,j]
        data.iloc[row_len,j]=sq_sum
        data.iloc[row_len+1,j]=sq_sum**0.5
    for j in range(1,col_len):
        for i in range(1,row_len-1):
            data.iloc[i,j]=data.iloc[i,j]*data.iloc[row_len+2,j]/data.iloc[row_len+1,j]
    #For generating max mean.
    for j in range(1,col_len):
        max=data.iloc[1,j]
        min=data.iloc[1,j]
        for i in range(2,row_len-2):
            if(max<data.iloc[i+1,j]):
                #print(" max ",max,"data.iloc[i+1,j]  ", data.iloc[i+1,j])
                max=data.iloc[i+1,j]
            if(min>data.iloc[i+1,j]):
                min=data.iloc[i+1,j]
            if(data.iloc[row_len-1,j]==1):
                data.iloc[row_len+3,j]=max     
                data.iloc[row_len+4,j]=min
            else:
                data.iloc[row_len+3,j]=min     
                data.iloc[row_len+4,j]=max
    print(data)
    s1=data.copy()
    s1[col_len]=np.NaN
    s1.iloc[0,col_len]='SQRT_SUM_S1'
    s2=s1.copy()
    s2.iloc[0,col_len]='SQRT_SUM_S2'
    #print("s1 is\n",s1)
    #print("s2 is\n",s2)
    for j in range(1,col_len):
        for i in range(1,row_len-1):
            s1.iloc[i,j]=(s1.iloc[i,j]-s1.iloc[row_len+3,j])**2
            s2.iloc[i,j]=(s2.iloc[i,j]-s2.iloc[row_len+4,j])**2
    #print("s1 modify is\n",s1)
    #print("s2 modify is\n",s2)
    for i in range(1,row_len-1):
        sum1=0
        sum2=0
        for j in range(1,col_len):
            sum1+=s1.iloc[i,j]
            sum2+=s2.iloc[i,j]
        s1.iloc[i,col_len]=sum1**0.5
        s2.iloc[i,col_len]=sum2**0.5
    print("s1 matrix with all parameters:\n",s1)
    print("s2 matrix with all parameters:\n",s2)
    topsis_score=s1.iloc[0:row_len-1,0:2]
    topsis_score.iloc[0,1]='SCORE'
    for i in range(1,row_len-1):
        topsis_score.iloc[i,1]=(s2.iloc[i,col_len]/(s2.iloc[i,col_len]+s1.iloc[i,col_len]))
    print("Final Topsis score\n",topsis_score)
    # ahp_in = pd.DataFrame(data)
    # print(data)
    topsis_score.to_excel('topsis_score.xlsx',index=False,header=False)


            
        
    #print(data)
if __name__=="__main__":
    topsis_calc()