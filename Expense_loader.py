import pandas as pd
from bs4 import BeautifulSoup
import subprocess

input_file = open("Config.txt","r")
input_file_lines = input_file.readlines()
year = int(input_file_lines[0].split(": ")[1])
topic_list = input_file_lines[1].split(": ")[1].split(", ")
input_file.close()

# expense sheet in expense tracker file
expensesSheetOriginal = pd.read_excel("Expense_tracker.ods",sheet_name="expenses")
# expense sheet in expense tracker file
summarySheetOriginal = pd.read_excel("Expense_tracker.ods",sheet_name="summary")

# month dictionary to number of days
month_dict = {"Jan":[1,31],"Feb":[2,31],"Mar":[3, 29] if (year%4 == 0) else [3, 28],"Apr":[4, 31],"May":[5, 30],"Jun":[6, 31],"Jul":[7, 30],"Aug":[8, 31],"Sep":[9, 31],"Oct":[10, 30],"Nov":[11, 31],"Dec":[12, 30]}

latest_day_index = 0
if(expensesSheetOriginal.empty):
    latest_day_index = 0
else:
    for i in month_dict:
        if(i != expensesSheetOriginal["Date"][0].split(" ")[0]):
            latest_day_index = latest_day_index + month_dict[i][1]
        else:
            break
    latest_day_index += int(expensesSheetOriginal["Date"][0].split(" ")[1])

with open("My Activity.html") as f:
    soup = BeautifulSoup(f, 'html.parser')

f = open("output.txt","w")
count = 1

Dates = []
Expenses = []
Amounts = []
Month_Year = []

for i in soup.find_all("div",attrs={"class":"outer-cell"}):
    content_tag = i.find_all("div",attrs={"class":"content-cell"})[0]
    transaction_tag = i.find_all("div",attrs={"class":"content-cell"})[2]
    content_tag.find("br").replace_with(' ; ')
    for i in transaction_tag.find_all("br"):
        i.replace_with(" ")
    if(transaction_tag.get_text().split(" ")[-2].split("\u2003")[-1] == "Completed"):
        # print(count, content_tag.get_text().split(" "))
        expense_amount = content_tag.get_text().split(" ")[1]
        if (("Paid" in content_tag.get_text().split(" ")) and ("to" in content_tag.get_text().split(" "))) or ("Sent" in content_tag.get_text().split(" ")):
            try:
                expense = content_tag.get_text().split(" ")[3:content_tag.get_text().split(" ").index("using")]
            except ValueError:
                expense = content_tag.get_text().split(" ")[3:content_tag.get_text().split(" ").index(";")]
            expense_amount = content_tag.get_text().split(" ")[1].split("₹")[1]
            expense_topic = ""
            for j in expense:
                expense_topic += j
                expense_topic += " "
            if(expense == []):
                expense_topic = "miscellaneous"
        else:
            expense_topic = "Recieved"
            expense_amount = "-" + content_tag.get_text().split(" ")[1].split("₹")[1]
        expense_date_list = content_tag.get_text().split(" ")[content_tag.get_text().split(" ").index(";")+1:content_tag.get_text().split(" ").index(";")+4]
        expense_date = ""
        for j in expense_date_list:
            expense_date = expense_date + j
            expense_date += " "
        expense_date = expense_date.replace(",","")
        current_day_index = 0
        for i in month_dict:
            if(i != expense_date.split(" ")[0]):
                current_day_index = current_day_index + month_dict[i][1]
            else:
                break
        current_day_index += int(expense_date.split(" ")[1])
        f.write(str(expense_topic))
        f.write(" ")
        f.write(str(expense_amount))
        f.write(" ")
        f.write(str(expense_date))
        f.write("\n")
        if(expensesSheetOriginal.empty) and (int(expense_date.split(' ')[2]) == year):
            f.write(str(expense_topic))
            f.write(" ")
            f.write(str(expense_amount))
            f.write(" ")
            f.write(str(expense_date))
            f.write("\n")
            Dates.append(expense_date)
            Expenses.append(expense_topic)
            Amounts.append(expense_amount)
            Month_Year.append(f"{month_dict[expense_date.split(" ")[0]][0]}-{expense_date.split(" ")[2]}")
            count += 1
        else:
            if((current_day_index > latest_day_index) and (int(expense_date.split(' ')[2]) == year)):
                f.write(str(expense_topic))
                f.write(" ")
                f.write(str(expense_amount))
                f.write(" ")
                f.write(str(expense_date))
                f.write("\n")
                Dates.append(expense_date)
                Expenses.append(expense_topic)
                Amounts.append(expense_amount)
                Month_Year.append(f"{month_dict[expense_date.split(" ")[0]][0]}-{expense_date.split(" ")[2]}")
                count += 1

df_delta = pd.DataFrame({"Date":Dates, "Expense":Expenses, "Amount":Amounts, "Month_Year":Month_Year})
expensesSheetOriginal = pd.concat([df_delta, expensesSheetOriginal])

with pd.ExcelWriter("Expense_tracker.ods",engine="odf") as writer:
    expensesSheetOriginal.to_excel(writer,sheet_name="expenses",index=False)
    summarySheetOriginal.to_excel(writer,sheet_name="summary",index=False)
f.close()

subprocess.run(["libreoffice",'Expense_tracker.ods'])