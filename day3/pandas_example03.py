import pandas as pd
list1 = [[1,2,3],[4,5,6],[7,8,9]]
data=pd.DataFrame(list1,columns=["col1","col2","col3"],index=["row1","row2","row3"])
print(data)
print("data col1 is: ",data["col1"])