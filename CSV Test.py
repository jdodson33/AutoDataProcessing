import pandas as pd
import matplotlib.pyplot as plt
import os


class JBITesting:
    def __init__(self):
        self.path = os.path.dirname(__file__)
        self.csv_file_path = self.path + "\\Data Files\\csv files"
        self.excel_path = self.path + "\\Data Files"

    def check_for_Excel(self):
        excel_doc = None
        for item in os.listdir(self.excel_path):
            if ".xlsx" in item:
                excel_doc = item

        if not excel_doc:
            print("No .xlsx documents in the appropriate folder. Please add one.")
            exit()
        else:
            return excel_doc

    def convert_sheets_to_csv(self):
        excel_doc = self.check_for_Excel()
        file_path = self.excel_path + "\\" + excel_doc

        all_sheets = pd.read_excel(file_path, sheet_name=None)
        sheets = all_sheets.keys()

        battery_tags = self.checkIn_check(file_path, sheets)

        for sheet in sheets:
            if sheet in ["Procedure", "Report", "Weights", "Data", "Check In"]:
                continue
            curr_sheet = pd.read_excel(file_path, sheet_name=sheet)
            curr_sheet.to_csv(self.csv_file_path + "\\%s.csv" % sheet, index=False)

        if not battery_tags:
            battery_tags = [n for n in range(1, len(sheets))]

        return battery_tags

    def checkIn_check(self, excel_path, sheets):
        tags = []
        for sheet in sheets:
            if "check" in str(sheet).lower():
                check_in_sheet = pd.read_excel(excel_path, sheet_name=sheet)
                check_in_sheet.to_csv(self.csv_file_path + "\\%s.csv" % sheet, index=False)
                check_in_df = pd.read_csv(self.csv_file_path + "\\Check In.csv")

                temp = check_in_df.iloc[:, 0]
                skip = 0
                for i in range(len(temp)):
                    if "battery" in str(temp[i]).lower():
                        skip = i + 1
                        break

                check_in_df = pd.read_csv(self.csv_file_path + "\\Check In.csv", skiprows=skip)
                check_in_df.dropna(how='all', inplace=True)

                for key in check_in_df.keys():
                    if "battery" in str(key).lower():
                        index = key
                        break

                for i in range(len(check_in_df.index)):
                    tags.append(check_in_df[index][i])

                print(tags)

        if os.path.exists(self.csv_file_path + "\\Check In.csv") and os.path.isfile(self.csv_file_path + "\\Check In.csv"):
            os.remove(self.csv_file_path + "\\Check In.csv")
        return tags

    def capture_data(self):
        batt_tags = self.convert_sheets_to_csv()

        total_RRC_data = []
        for workbook in os.listdir(self.csv_file_path):
            df = pd.read_csv(self.csv_file_path + "\\" + str(workbook))
            temp = df.iloc[:, 0]

            skip = 0
            for i in range(len(temp)):
                if str(temp[i]).lower() == "total time, s":
                    skip = i + 1
                    break

            df = pd.read_csv(self.csv_file_path + "\\" + str(workbook), skiprows=skip)
            df.dropna(how='all', inplace=True)  # clean the data - remove NaN

            RRC_data = []
            # print(df.keys())
            for i in range(len(df.index)):
                if df.iloc[i]["Data Acquisition Flag"] == "S" and df.iloc[i]["Mode"] == "DCHG":
                    RRC_data.append(df.iloc[i]["Step Time, S"] / 60)
                    # print(df.iloc[i]["Step Time, S"], df.iloc[i]["Data Acquisition Flag"], df.iloc[i]["Mode"])
            total_RRC_data.append(RRC_data)
            # print(total_RRC_data)
        return batt_tags, total_RRC_data

    def plot_data(self):
        batt_RC_spec = input("Please input the RC spec for the batteries being tested: ")
        batt_tags, total_RRC_data = self.capture_data()

        max_len = 0
        for num, row in enumerate(total_RRC_data):
            x_vals = [i for i in range(1, len(row) + 1)]
            batt_name = num
            if batt_tags and num < len(batt_tags):
                batt_name = batt_tags[num]

            plt.plot(x_vals, row, marker="o", label="RRC-%s" % batt_name)
            max_len = max(max_len, len(x_vals))
        # print(total_RRC_data)
        x = [0, max_len]
        y = [int(batt_RC_spec), int(batt_RC_spec)]
        plt.plot(x, y, label="Spec")

        plt.title("RRC Testing Data")
        plt.xlabel("Cycles")
        plt.ylabel("Reserve Capacity, mins")
        plt.legend()

        plt.show()


test = JBITesting()
test.plot_data()