import pandas as pd

#Reading the Excel file.
pm = pd.read_excel("pm.xlsx", sheet_name="main")
pms = pd.read_excel("pm.xlsx", sheet_name="state")

#Asking user to enter the name for the final Excel file. 
file_name = str(input("Enter final file Name with .xlsx at the end: "))

#Sorting the data.
sort = pm.sort_values(["Name", "Years in Office"], ascending=[True, False])

sort.to_excel(file_name, sheet_name="Sort", index=False)

#Filtering the Data.
filtered_data = pm.loc[(pm["Age"] > 75) & (pm["Years in Office"] > 5)]

#Adding the new sheet in the existing file and pasting the data.
with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
   filtered_data.to_excel(writer, sheet_name="Filter", index=False)

#Doing the Vlookup.
vlookup = pd.merge(pm, pms, on="Name", how="left")
vlookup["State"] = vlookup["State"].fillna("N/A")

with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
   vlookup.to_excel(writer, sheet_name="Vlookup", index=False)

#Adding Pivot table.
pt = vlookup.pivot_table(index="State", values="Name", aggfunc="count").reset_index()
pt.columns = ["state", "count of Names"]

with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
   pt.to_excel(writer, sheet_name="Pivot Table", index=False)

print("File is created")
