import openpyxl.utils
import openpyxl
import numbers


##define open_worksheet
#open the file, read it, return first worksheet

def open_worksheet(filename):
    income_excel = openpyxl.load_workbook(filename)
    data_sheet = income_excel.active
    return data_sheet

##define main function//
#call open_worksheet, call should_get_losses, call process_data

def main():
    data_sheet = open_worksheet("countyPopChange2020-2021.xlsx")
    get_losses = should_get_losses()
    process_data(data_sheet, get_losses)

##define should_get_losses
#ask user if we should get counties that LOST population

def should_get_losses():
    response = input("Should we get the counties that LOST population?")
    if response.lower() == "yes":
        return True
    else:
        return False

##define process_data
#calculate the percentage change
#print the counties that gained/lost based on user input

def process_data(data_sheet, should_get_losses):

    percentage_population_changes = []
    data_sheet.delete_rows(0)
    for row in data_sheet.rows:
        first_cell = row[9]
        second_cell = row[11]
        newindex = (int(second_cell.value)/int(first_cell.value))*100
        percentage_population_changes.append(newindex)

        if (should_get_losses):
            if (abs(newindex) > 2): print(str(row[5].value) + ", " + str(row[6].value) + " " + str(newindex) + "%")
        else:
            if (newindex > 1.5): print(str(row[5].value) + ", " + str(row[6].value) + " " + str(newindex) + "%")

main()

