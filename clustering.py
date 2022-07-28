import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.preprocessing import MinMaxScaler

df = pd.read_excel("./Files/Output/Result.xlsx", sheet_name="Reports")

# Keep actual data, remove forecasted
df = df[df["Remark"] == "Actual Date"]
# Keep latest record for each project
df = df.drop_duplicates(subset=["Project ID"], keep="last")
# Keep Project ID, VAC, VAC(t)
df = df[["Project ID", "VAC", "VAC(t)"]]

# Normalisation
scaler = MinMaxScaler()
scaler.fit(df[["VAC", "VAC(t)"]])
df[["VAC(n)", "VAC(t)(n)"]] = scaler.transform(df[["VAC", "VAC(t)"]])

km = KMeans(n_clusters=3)
y_predicted = km.fit_predict(df[["VAC(n)", "VAC(t)(n)"]])

df["Cluster"] = y_predicted

df1 = df[df["Cluster"] == 0]
df2 = df[df["Cluster"] == 1]
df3 = df[df["Cluster"] == 2]
plt.scatter(df1["VAC(n)"], df1["VAC(t)(n)"])
plt.scatter(df2["VAC(n)"], df2["VAC(t)(n)"])
plt.scatter(df3["VAC(n)"], df3["VAC(t)(n)"])
plt.show()