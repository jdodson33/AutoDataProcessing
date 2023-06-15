import pandas as pd
import matplotlib.pyplot as plt
import os

csv_file_path = os.path.dirname(__file__)
csv_file_path += "\\Data Files\\csv files"

excel_path = os.path.dirname(__file__) + "\\Data Files"
excel_doc = None
for item in os.listdir(excel_path):
    if ".xlsx" in item:
        excel_doc = item

if not excel_doc:
    print("No .xlsx documents in the appropriate folder")
    exit()

file_path = excel_path + "\\" + excel_doc
all_sheets = pd.read_excel(file_path, sheet_name=None)
sheets = all_sheets.keys()

for sheet in sheets:
    if sheet in ["Procedure", "Report", "Weights", "Data", "Check In"]:
        continue
    curr_sheet = pd.read_excel(file_path, sheet_name=sheet)
    curr_sheet.to_csv(csv_file_path + "\\%s.csv" % sheet, index=False)

total_RRC_data = []
for workbook in os.listdir(csv_file_path):
    print(workbook)
    df = pd.read_csv(csv_file_path + "\\" + str(workbook))
    temp = df.iloc[:, 0]

    # skip header to the actual data
    skip = 0
    for i in range(len(temp)):
        if str(temp[i]).lower() == "total time, s":
            skip = i + 1
            break

    df = pd.read_csv(csv_file_path + "\\" + str(workbook), skiprows=skip)
    df.dropna(how='all', inplace=True)  # clean the data - remove NaN

    # Find and add RC time for each cycle to a list
    RRC_data = []
    for i in range(len(df.index)):
        if df.iloc[i]["Data Acquisition Flag"] == "S" and df.iloc[i]["Mode"] == "DCHG":
            RRC_data.append(df.iloc[i]["Step Time, S"] / 60)
    total_RRC_data.append(RRC_data)


max_len = 0
for num, row in enumerate(total_RRC_data):
    x_vals = [i for i in range(1, len(row)+1)]
    num += 1
    plt.plot(x_vals, row, marker="o", label="RRC-%s" % num)
    max_len = max(max_len, len(x_vals))

x = [0, max_len]
y = [70, 70]
plt.plot(x, y, label="Spec")

plt.title("RRC Testing Data")
plt.xlabel("Cycles")
plt.ylabel("Reserve Capacity, mins")
plt.legend()

plt.show()
