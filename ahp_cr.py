import numpy as np
import sys
import pandas as pd
def ahp_cr():
    d1=pd.read_excel("ahp_in.xlsx",header=None)
    #print(d1)
    wt=pd.read_excel("ahp_weights.xlsx",header=None)
    #print(wt)
    d1.loc[len(d1),:] = [np.NAN for i in range(len(d1.columns))]
    #print(d1)
    col_len=len(d1.columns)
    row_len=len(d1)
    d1.iloc[row_len-1,0]='sum'
    for j in range(1,col_len):
        sum=0
        for i in range(1,row_len-1):
            sum+=d1.iloc[i,j]
        d1.iloc[row_len-1,j]=sum
    #print(d1)
    max_eigen=0
    for i in range(1,row_len-1):
        max_eigen+=d1.iloc[row_len-1,i]*wt.iloc[i,1]
    print("Maximum eigen value is :",max_eigen)
    ci=(max_eigen-(row_len-2))/((row_len-2)-1)
    print("The CI is :",ci)
    ri_data=pd.read_excel("ri_table.xlsx",header=None)
    #Finding out CR
    for i in range(1,29):
        if(ri_data.iloc[i,0]==row_len-2):
            ri=ri_data.iloc[i,1]
        else:
            continue
    print("For ",row_len-2," order matrix RI is :",ri)
    cr=ci/ri
    if(cr<=0.1):
        print("AHP table acceptable for CR value :",cr)
    else:
        print("Redo AHP table,You CR value = ",cr,".\nIt must be less than 0.1.")
if __name__=="__main__":
    ahp_cr()