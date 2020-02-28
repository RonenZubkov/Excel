# Writing to an excel
# sheet using Python
import xlrd
from collections import Counter
import xlsxwriter


def main():
    """"""
    create_new_sheet()
# ----------------------------------------------------------------------

def get_data_from_xl():

    # Give the location of the file
    loc = ("טופס הזמנה פברואר.xlsx")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    # sheet.cell_value(0, 0)
    # print(sheet.cell_value(0, 0))
    return tuple(sheet.col(6))


def get_data():
    cats_food = []
    catFoodList = get_data_from_xl()

    for num, catFood in enumerate(catFoodList):
        # value = sheet.cell(num, 6).value
        # value = catFoodList[num]
        if (catFoodList[num]):
            cats_food.append(catFoodList[num].value)

    return cats_food


def count_food():
    cats_food = get_data()
    return Counter(sorted(cats_food))


def create_new_sheet():
    workbook = xlsxwriter.Workbook('Counted.xlsx')
    worksheet = workbook.add_worksheet()
    sorted_food_cat = count_food()
    headers = ['שם פריט', 'כמות']
    column = 0

    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # rewrite this shit as soon as possible #
    for row, item in enumerate(sorted_food_cat):
        worksheet.write(row + 1, column, item)
        # worksheet.write(row + 1, column, item)
        worksheet.write(row + 1, column + 1, sorted_food_cat[item])

    workbook.close()


if __name__ == "__main__":
    main()
