import pandas as pd

df = pd.read_csv('air_travel.csv')
print(df.info())
print(df.describe())
print(df['Destination'].value_counts())

#print(df.to_string())