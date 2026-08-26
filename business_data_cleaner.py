import pandas as pd

df = pd.read_excel('data/messy_sales.xlsx')

print(df)

df.info()

print(df.isnull().sum())


print(df.duplicated().sum())


print(df[df.duplicated()])


df=df.drop_duplicates()

print(df[df.duplicated()])


print(len(df))

print(df["Customer"].unique())


df["Customer"]=df["Customer"].str.lower().str.strip()


print(df["Customer"].unique())

df["Product"]=df["Product"].str.lower().str.strip()

print(df["Product"].unique())

print(df[["Customer","Product"]].duplicated().sum())
print(df[df[["Customer","Product"]].duplicated(keep=False)])


df["Date"]=pd.to_datetime(df["Date"],format="mixed", dayfirst=True)

print(df["Date"])


print(df.duplicated().sum())

print(df[df.duplicated()])

df=df.drop_duplicates()
print(df.duplicated().sum())
print(len(df))


print(df.isnull().sum())
print(df[df.isnull().any(axis=1)])

df["Missing_Date"] = df["Date"].isnull()

print(df[["Customer","Product","Date","Missing_Date"]])


df["Missing_Quantity"] = df["Quantity"].isnull()
df["Missing_Customer"] = df["Customer"].isnull()
df["Invalid_Quantity"] = df["Quantity"] <= 0 

df["Quantity"] = df["Quantity"].fillna(1)
df["Customer"] = df["Customer"].fillna("Unknown")
print(df[["Customer", "Quantity", "Price"]])

df["Invalid_Quantity"] = df["Quantity"] <= 0

print(df[df["Invalid_Quantity"]])

df["Price"] = df["Price"].fillna( 
    df.groupby("Product")["Price"].transform("first") )
print(df[["Customer",
          "Product",
          "Quantity",
          "Price",
          "Missing_Customer",
          "Missing_Quantity"
          ]])


print(df.isnull().sum())

df["Valid_Sale"] = ~df["Invalid_Quantity"]

print(df[["Customer", "Product", "Quantity", "Invalid_Quantity", "Valid_Sale"]])

print(df[["Customer", "Product", "Date"]].sort_values("Date"))

df["Suspicious_Date"] = (
    (df["Date"].dt.month == 12) &
    (df["Date"].dt.year == 2024)
)

print(df[df["Suspicious_Date"]])


print(df.duplicated().sum())

df = df.drop_duplicates()

print(df.duplicated().sum())
print(len(df))


df.loc[
    (df["Customer"] == "ramesh") &
    (df["Product"] == "laptop"),
    "Date"
] = pd.Timestamp("2024-05-12")

df["Suspicious_Date"] = (
    (df["Date"].dt.month == 12) &
    (df["Date"].dt.year == 2024)
)

print(df[df["Suspicious_Date"]])

df["Total_Sales"] = df["Quantity"] * df["Price"]

print(df[[
    "Customer",
    "Product",
    "Quantity",
    "Price",
    "Total_Sales"
]])


df["Valid_Revenue"] = df["Total_Sales"].where(
    df["Valid_Sale"],
    0
)

print(df[[
    "Customer",
    "Product",
    "Quantity",
    "Total_Sales",
    "Valid_Sale",
    "Valid_Revenue"
]])

total_revenue = df["Valid_Revenue"].sum()

print("Total Valid Revenue:", total_revenue)

product_revenue = df.groupby("Product")["Valid_Revenue"].sum()
product_revenue = product_revenue.sort_values(ascending=False)
print(product_revenue)

best_product = product_revenue.idxmax()

print("Best Product:", best_product)

customer_revenue = df.groupby("Customer")["Valid_Revenue"].sum()

customer_revenue = customer_revenue.sort_values(ascending=False)

print(customer_revenue)

best_customer = customer_revenue.idxmax()

print("Best Customer:", best_customer)

print("\n========== SALES REPORT ==========")

print("Total Revenue: ₹", total_revenue)

print("Best Product:", best_product)

print("Best Product Revenue: ₹", product_revenue.max())

print("Best Customer:", best_customer)

print("Best Customer Revenue: ₹", customer_revenue.max())

print("Invalid Sales:", df["Invalid_Quantity"].sum())

print("Unknown Customer Revenue: ₹",
      df.loc[df["Customer"] == "Unknown", "Valid_Revenue"].sum())

print("==================================")