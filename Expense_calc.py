import subprocess
import pandas as pd

input_file = open("Config.txt","r")
input_file_lines = input_file.readlines()
year = int(input_file_lines[0].split(": ")[1])
topic_list = input_file_lines[1].split(": ")[1].split(", ")
input_file.close()

# expense sheet in expense tracker file
expensesSheetOriginal = pd.read_excel("Expense_tracker.ods",sheet_name="expenses")
# expense sheet in expense tracker file
summarySheetOriginal = pd.read_excel("Expense_tracker.ods",sheet_name="summary")

delta_summary = {
    "Category": topic_list + [" ", "Total"]
}

for i in range(12):
    delta_summary[f"{i+1}-{year}"] = [0 for topic_list_index in range(len(topic_list))] + [" ", " "]
    for j in range(expensesSheetOriginal.shape[0]):
        if(expensesSheetOriginal.iloc[j]["Month_Year"] == f"{i+1}-{year}"):
            delta_summary[f"{i+1}-{year}"][topic_list.index(expensesSheetOriginal.iloc[j]["Notes"])] += expensesSheetOriginal.iloc[j]["Amount"]
    delta_summary[f"{i+1}-{year}"][-1] = sum(delta_summary[f"{i+1}-{year}"][0:len(topic_list)])

# print(expensesSheetOriginal.iloc[0]["Amount"])

df_delta_summary = pd.DataFrame(delta_summary)

with pd.ExcelWriter("Expense_tracker.ods",engine="odf") as writer:
    expensesSheetOriginal.to_excel(writer,sheet_name="expenses",index=False)
    df_delta_summary.to_excel(writer,sheet_name="summary",index=False)

subprocess.run(["libreoffice",'Expense_tracker.ods'])