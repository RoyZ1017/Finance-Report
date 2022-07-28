import xlsxwriter
import pandas as pd
from matplotlib import pyplot as plt
import numpy as np


# class which contains functions to manipulate excel documents
class ExcelUtility:
    def __init__(self, excel_file_name):
        self.excel_file_name = excel_file_name

    # creates a new excel workbook and worksheet
    # returns the workbook and worksheet in the form of a tuple --> (workbook, worksheet)
    def open_excel(self, sheet_name):
        try:
            workbook = xlsxwriter.Workbook(self.excel_file_name)
            worksheet = workbook.add_worksheet(sheet_name)
            return workbook, worksheet
        except Exception as e:
            print(e)
            return None

    # creates a new excel sheet
    def create_new_sheet(self, workbook, sheet_name):
        worksheet = workbook.add_worksheet(sheet_name)
        return worksheet

    # write message in specific cell
    def write_in_cell(self, row, col, message, worksheet, format={}):  # message can be in for of str or list
        try:
            worksheet.write(row, col, message, format)
            return message
        except Exception as e:
            print(e)
            return None

    # insert pandas dataframe into excel document
    # returns the writer used to insert the datafame into excel
    def convert_data_frame_to_xlsx(self, df, sheet, row=0, col=0, writer=None):
        try:
            if not writer:
                writer = pd.ExcelWriter(self.excel_file_name, engine='xlsxwriter', date_format='mmmm dd yyyy')
            df.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=row, startcol=col)
            return writer
        except Exception as e:
            print(e)
            return None

    # insert a png into excel document
    def insert_image(self, row, col, img_name, worksheet):
        try:
            worksheet.insert_image(row, col, img_name)
            return True
        except Exception as e:
            print(e)
            return None

    # change the cell formatting e.g. money, text colour, center allignment etc.
    def cell_format(self, workbook, formatting):
        try:
            format = workbook.add_format(formatting)
            return format
        except Exception as e:
            print(e)
            return None

    # merge multiple cells together
    def merge_cells(self, workbook, worksheet, message, m_range, format={'bold': 1, 'border': 1, 'align': 'center',
                                                                'valign': 'vcenter'}):
        try:
            merge_format = workbook.add_format(format)
            worksheet.merge_range(m_range, message, merge_format)
        except Exception as e:
            print(e)
            return None

    # close and saves changes made to an excel workbook
    @classmethod
    def close_excel(cls, workbook):
        try:
            workbook.close()
        except Exception as e:
            print(e)
            return None


# creates a new excel workbook
excel = ExcelUtility("Finance Report.xlsx")
excel.open_excel("Sheet1")


# define variables
revenue = {"Item": [], "Price": [], "Date":[]}
expense = {"Item": [], "Price": [], "Date": []}
all_expendatures = {"Item": [], "Price": [], "Date": []}
needs = {"Item": [], "Price": [], "Date": []}
wants = {"Item": [], "Price": [], "Date": []}
others = {"Item": [], "Price": [], "Date": []}
total_revenue = 0
total_expenses = 0
total_needs = 0
total_wants = 0
total_others = 0

# gather user input
while True:
    new_item = input().split(" ")
    if new_item == ["result"]:  # break out of loop and return and fill out the excel spreadsheet
        break
    else:
        try:
            # if the item is classified as revenue insert the information into the revenue dictionary
            if new_item[2] == "r":
                revenue["Item"].append(new_item[0])
                revenue["Price"].append(float("{:.2f}".format(float(new_item[1]))))  # all prices are rounded to 2 floating points of precision
                revenue["Date"].append(new_item[3])
                total_revenue += float("{:.2f}".format(float(new_item[1])))  # increment total_revenue

            # if the item is classified as expense insert the information into the expense dictionary
            elif new_item[2] == "e":
                # all_expendatures keeps track of all items inputted including duplicates
                all_expendatures["Item"].append(new_item[0])
                all_expendatures["Price"].append(float("{:.2f}".format(float(new_item[1]))))
                all_expendatures["Date"].append(new_item[4])

                # expense drops all duplicates and instead keeps track of total spending on a item
                if new_item[0] in expense["Item"]:
                    index = expense["Item"].index(new_item[0])
                    expense["Price"][index] += float("{:.2f}".format(float(new_item[1])))
                else:
                    expense["Item"].append(new_item[0])
                    expense["Price"].append(float("{:.2f}".format(float(new_item[1]))))
                    expense["Date"].append(new_item[4])
                total_expenses += float("{:.2f}".format(float(new_item[1])))

                # if the item is classified as a need insert the information into the need dictionary
                if new_item[3] == "n":
                    needs["Item"].append(new_item[0])
                    needs["Price"].append(float("{:.2f}".format(float(new_item[1]))))
                    needs["Date"].append(new_item[4])
                    total_needs += float("{:.2f}".format(float(new_item[1])))

                # if the item is classified as a want insert the information into the want dictionary
                elif new_item[3] == "w":
                    wants["Item"].append(new_item[0])
                    wants["Price"].append(float("{:.2f}".format(float(new_item[1]))))
                    wants["Date"].append(new_item[4])
                    total_wants += float("{:.2f}".format(float(new_item[1])))

                # if the item is classified as other insert the information into the other dictionary
                elif new_item[3] == "o":
                    others["Item"].append(new_item[0])
                    others["Price"].append(float("{:.2f}".format(float(new_item[1]))))
                    others["Date"].append(new_item[4])
                    total_others += float("{:.2f}".format(float(new_item[1])))
                # if the item is neither classified as a need, want, or other raise an exception as it is an invalid input
                else:
                    raise Exception
            # if the item is neither classified as revenue or expense raise an exception as it is an invalid input
            else:
                raise Exception

        # if the input is invalid prompt the user to try again
        except Exception as e:
            print("invalid input")
            print("please ensure that the input format is --> item name, amount, "
                  "r/e(revenue or expense), n/w/o (need, want or other if the item is an expense), "
                  "date (mm/dd/yyyy)")



# calculate the net income
net_income = total_revenue - total_expenses

# information displayed on sheet1 of the excel file
# shee1 contains an income statement
# turns the revenue and expense into a pandas dataframe
revenue_df = pd.DataFrame(revenue, columns=["Item", "Price"]).sort_values("Price", ascending=False)
expense_df = pd.DataFrame(expense, columns=["Item", "Price"]).sort_values("Price", ascending=False)
# empty dataframe is used to add blank lines in the dataframe seperating the revenue and expense section
empty_df = pd.DataFrame({"Item": [None] * 3, "Price": [None] * 3}, columns=["Item", "Price"])
# joins the revenue and expense dataframe into one big dataframe
revenue_and_expense = pd.concat([revenue_df, empty_df, expense_df])

start_index = 4
# inserts the revenue and expense dataframe into the excel document on Sheet1
writer = excel.convert_data_frame_to_xlsx(revenue_and_expense, "Sheet1", start_index, 0)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# sets some excel cell formats so that the number inserted is displayed in accounting form
total = excel.cell_format(workbook, {"num_format": '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)', "underline": True})
money = excel.cell_format(workbook, {"num_format": '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'})
worksheet.set_column(1, 1, None, money)

# create a header for Sheet1 using merge
excel.merge_cells(workbook, worksheet, "Income Statement", "A2:D2")

# adds information such as total revenue, total expense, and net income into the excel
excel.write_in_cell(start_index - 1, 0, "Revenue", worksheet)
excel.write_in_cell(start_index - 1, 1, "Amount", worksheet)
excel.write_in_cell(revenue_df.shape[0] + start_index, 0, "Total Revenue", worksheet)
excel.write_in_cell(revenue_df.shape[0] + start_index, revenue_df.shape[1], total_revenue, worksheet, total)

excel.write_in_cell(revenue_df.shape[0] + start_index + 2, 0, "Expenses", worksheet)
excel.write_in_cell(revenue_and_expense.shape[0] + start_index, 0, "Total Expenses", worksheet)
excel.write_in_cell(revenue_and_expense.shape[0] + start_index, revenue_and_expense.shape[1], total_expenses, worksheet, total)

excel.write_in_cell(revenue_and_expense.shape[0] + start_index + 1, 0, "Net Income", worksheet)
excel.write_in_cell(revenue_and_expense.shape[0] + start_index + 1, revenue_and_expense.shape[1], net_income, worksheet, total)

# information displayed on sheet2 of the excel file
# sheet2 displays the seperation of needs wants and others and displays various graphs
# turns needs, wants, and others into a pandas dataframe
needs_df = pd.DataFrame(needs, columns=["Item", "Price", "Date"])
wants_df = pd.DataFrame(wants, columns=["Item", "Price", "Date"])
others_df = pd.DataFrame(others, columns=["Item", "Price", "Date"])
empty_df = pd.DataFrame({"Item": [None] * 3, "Price": [None] * 3, "Date": [None] * 3}, columns=["Item", "Price", "Date"])
# joins needs, wants, and others into one big dataframe
needs_wants_others = pd.concat([needs_df, empty_df, wants_df, empty_df, others_df])

# insert the dataframe into the excel document on Sheet2
excel.convert_data_frame_to_xlsx(needs_wants_others, "Sheet2", start_index, 0, writer)

# create a second worksheet called "Sheet2"
worksheet2 = writer.sheets["Sheet2"]
worksheet2.set_column(1, 1, None, money)

# create a header for Sheet2 using merge
excel.merge_cells(workbook, worksheet2, "Needs and Wants", "A2:D2")
# insert information such as total needs, total wants, and total others into the excel
excel.write_in_cell(start_index - 1, 0, "Needs", worksheet2)
excel.write_in_cell(start_index - 1, 1, "Amount", worksheet2)
excel.write_in_cell(start_index - 1, 2, "Date", worksheet2)

excel.write_in_cell(needs_df.shape[0] + start_index, 0, "Total Needs", worksheet2)
excel.write_in_cell(needs_df.shape[0] + start_index, needs_df.shape[1], total_needs, worksheet2, total)

excel.write_in_cell(needs_df.shape[0] + start_index + empty_df.shape[0] - 1, 0, "Wants", worksheet2)
excel.write_in_cell(needs_df.shape[0] + wants_df.shape[0] + empty_df.shape[0] + start_index, 0, "Total Wants", worksheet2)
excel.write_in_cell(needs_df.shape[0] + wants_df.shape[0] + empty_df.shape[0] + start_index, wants_df.shape[1], total_wants, worksheet2, total)

excel.write_in_cell(needs_df.shape[0] + wants_df.shape[0] + start_index + 2 * (empty_df.shape[0]) - 1, 0, "Others", worksheet2)
excel.write_in_cell(needs_df.shape[0] + wants_df.shape[0] + others_df.shape[0] + 2 * empty_df.shape[0] + start_index, 0, "Total Others", worksheet2)
excel.write_in_cell(needs_df.shape[0] + wants_df.shape[0] + others_df.shape[0] + 2 * empty_df.shape[0] + start_index, others_df.shape[1], total_others, worksheet2, total)

# create scatter plot showing items bought and the day the were bought and saves it as a png
all_expendatures_df = pd.DataFrame(all_expendatures, columns=["Item", "Price", "Date"]).sort_values("Date")
all_expendatures_df.plot(x="Date", y="Price", kind="scatter")
plt.xlabel('Date')
plt.ylabel('Price')
plt.title("Expendatures")
scatter_plot = plt.gcf()
scatter_plot.set_size_inches(6, 4)
scatter_file_name = "scatter.png"
plt.savefig(scatter_file_name)
plt.show()

# create pie chart showing the current spending distrubution between items and saves it as a png
plt.pie(expense_df.loc[:,"Price"], labels=expense_df.loc[:,"Item"], autopct='%1.1f%%')
plt.legend(expense_df.loc[:,"Item"])
plt.title("Spending Distribution")
pie_chart = plt.gcf()
pie_chart.set_size_inches(6, 4)
pie_chart_file_name = "pie chart.png"
plt.savefig(pie_chart_file_name)
plt.show()

# create double bar graph comparing the user's current spending distribution and the reconmmended spending distribution and saves it as a png
current_spending = [total_needs, total_wants, total_others, net_income]
# recommended spending is a 50/30/20 split where 50% of your income goes to needs, 30% goes to wants, and 20% goes to savings
recomended_spending = [total_revenue * 0.5, total_revenue * 0.3, 0, total_revenue * 0.2]
x = ["Needs", "Wants", "Others", "Potential Savings"]
x_axis = np.arange(len(x))
plt.bar(x_axis - 0.2, current_spending, width=0.4, label="Current Spending Distribution")
plt.bar(x_axis + 0.2, recomended_spending, width=0.4, label="Recomended Spending Distribution")
plt.xticks(x_axis, x)
plt.title("Current Spending Distribution VS Recomended Spending Distribution")
plt.ylabel("Amount")
plt.legend()
double_bar_graph = plt.gcf()
double_bar_graph.set_size_inches(6, 4)
double_bar_graph_file_name = "double bar graph.png"
plt.savefig(double_bar_graph_file_name)
plt.show()

# insert graphs into excel
image_width = 10
image_height = 20
excel.insert_image(0, needs_wants_others.shape[1] + 2, scatter_file_name, worksheet2)  # insert scatter plot
excel.insert_image(0, needs_wants_others.shape[1] + image_width + 2, pie_chart_file_name, worksheet2)  # insert pie chart
excel.insert_image(image_height, needs_wants_others.shape[1] + 2, double_bar_graph_file_name, worksheet2)  # insert double bar graph

# close excel workbook and saves the file
excel.close_excel(writer.book)
