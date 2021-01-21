import pandas as pd
def removeUnameColumns(data):
    data.dropna(inplace=True,how='all')
    columns=data.columns.tolist()
    for i in columns:
        if(i.startswith('Unnamed')):
            columns.remove(i)
    return pd.DataFrame(data,columns=columns)