import numpy as np
import sys
import pandas as pd

# Get the number of rows and columns from the user

# Create a list to store the user input for the data
data = []


def number_check(val):
    """
    val: an object of type string
    function checks whether a string has number value or a if it's just a string
    return: returns False if it not a number, returns the float value if it
    is a number
    """
    for a in val:
        if (ord(a) >= 65 and ord(a) <= 90) or (ord(a) >= 97 and ord(a) <= 122):
            return False
    return float(val)


def ahp_in():
    # Loop through the rows and columns to get the user input for the data
    print("To find out weights of given criterias using AHP")
    print("The number of criterias must be between 3 and 30")
    num = int(input("Enter the number of criterias"))
    if num > 2 and num < 30:
        num += 1
        data = pd.DataFrame(np.NAN, index=range(num), columns=range(num))
        # print(data)
        for i in range(num):
            for j in range(num):
                if i == 0:
                    if j != 0:
                        print("Enter the", j, "th criteria ")
                        value = input()
                        data.iloc[0, j] = value
                        data.iloc[j, 0] = value
                    else:
                        data.iloc[i, i] = "criteria"
                else:
                    if j == i:
                        data.iloc[i, j] = 1.0
                    elif i < j:
                        print(
                            "Weight of", data.iloc[i, 0], " / ", data.iloc[0, j], ": "
                        )
                        value = input()
                        check_val = number_check(value)
                        if not check_val:
                            data.iloc[i, j] = value
                            data.iloc[j, i] = value
                        else:
                            data.iloc[i, j] = check_val
                            data.iloc[j, i] = 1 / check_val
                    else:
                        continue
        ahp_in = pd.DataFrame(data)
        print(data)
        ahp_in.to_excel("ahp_in4.xlsx", index=False, header=False)
    else:
        print("Enter proper criteria and try again.")


if __name__ == "__main__":
    ahp_in()
