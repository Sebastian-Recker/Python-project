
import pandas as pd

Excel_file = pd.ExcelFile("Verkäufe_je_Artikel.xlsx")
print(Excel_file.sheet_names)
df = Excel_file.parse("Verkäufe je Artikel und Jahr")

def margin_if_neg(row):
    if row["VK Preis Brutto 2014"]-row["EK Preis"]<0:
        return (row["VK Preis Brutto 2014"]-row["EK Preis"])*row["Sales Qty. 2014"]

df["neg_margin"] = df.apply(lambda row: margin_if_neg(row),axis=1)

print(df)
print(df["neg_margin"].sum(axis=0))
df.to_csv("test.csv")
