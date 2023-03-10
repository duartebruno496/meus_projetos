import io 
import requests 
import matplotlib.pyplot as plt 
import pandas as pd 
import gzip
import seaborn as sns 

url = "http://data.insideairbnb.com/brazil/rj/rio-de-janeiro/2020-0-1-21/data/listings.csv.gz"

response = requests.get(url)
bytes_io = io.BytesIO(response.content)
with gzip.open(bytes_io, 'rt') as read_file:
    ds = pd.read_csv(read_file)

dim = ds.shape 
print(dim)
print(ds.columns)

def remove_simbolo_monetario(x):
    a = x [1:]
    result = ""
    for i in a:
        if i.isdigit() is True:
            result = result + i
        return result
ds["price"] = pd.to_numeric(ds["price"].apply(remove_simbolo_monetario), errors= "ignore")  


plt.figure(figsize = (7,7))
sns.set(style = "whitegrid")
f = sns.distplot(ds["price"])

