import numpy as np
import sys
import pandas as pd

eval_mat = []

def number_check(val):

    for a in val:
        if (ord(a) >= 65 and ord(a) <= 90) or (ord(a) >= 97 and ord(a) <= 122):
            return False
    return float(val)

def topsis_in():
    print("To decide multi criterias using TOPSIS.")
    print("The number of criterias must be between 3 and 30.")
    cri_no = int(input("Enter the number of criterias:"))
    wt=pd.read_excel("ahp_weights.xlsx",header=None)
    row_len=len(wt)
    if(cri_no>2 and cri_no<30):
        if(cri_no==row_len):
            cri_no+=1
            alt_no = int(input("Enter the number of alternatives:"))
            alt_no+=1
            eval_mat = pd.DataFrame(np.NAN, index=range(alt_no+1), columns=range(cri_no))
            eval_mat.iloc[0,0]='criteria'
            eval_mat.iloc[alt_no,0]='COST/BENEFIT'
            for i in range(1,alt_no):
                print("Enter alternative ",i," : ")
                value = input()
                eval_mat.iloc[i,0]=value 
            for j in range(1,cri_no):
                print("Enter criteria ",j," : ")
                value = input()
                eval_mat.iloc[0,j]=value
                print("If this Criteria is benefit enter 1,if cost enter 0") 
                value = input()
                check_val = number_check(value)
                if not check_val:
                    eval_mat.iloc[alt_no,j]=value
                else:
                    eval_mat.iloc[alt_no,j]=check_val
            for i in range(1,alt_no):
                print("For Alternative-> ",eval_mat.iloc[i,0],"\nEnter the value of the following criterias:\n==================\n")
                for j in range(1,cri_no):
                    print(i,j,">",eval_mat.iloc[0,j]," :")
                    value = input()
                    check_val = number_check(value)
                    if not check_val:
                        eval_mat.iloc[i,j]=value
                    else:
                        eval_mat.iloc[i,j]=check_val
            print(eval_mat)
            topsis_in = pd.DataFrame(eval_mat)
            print(topsis_in)
            topsis_in.to_excel('topsis_in.xlsx',index=False,header=False)
        else:
            print("The criteria of AHP must be equal to criteria of TOPSIS.")
    else:
        print("Choose proper number of criterias.Restart.")
if __name__=="__main__":
    topsis_in()