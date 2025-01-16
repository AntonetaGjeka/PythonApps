import pandas as pd
#Load the file into the dataframe
df = pd.read_excel('assets.xlsx')
#Filter the dataframe accordingly
df = df[(df['Item']=='Laptop') & (df['Legal Entity']=='EDD') & (df['Asset Status']=='End of Life')]
df.to_excel('filtered_data.xlsx', sheet_name='Filtered Data from Assets')
