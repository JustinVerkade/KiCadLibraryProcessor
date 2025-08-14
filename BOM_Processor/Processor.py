from openpyxl.styles import Alignment, numbers
import openpyxl as pyxl
import datetime as dt
import pandas as pd
from tkinter.filedialog import askopenfilename

# USER CONSTANTS

TARGET_CSV = askopenfilename()

# LIBRARY CONSTANTS

LIBRARY_WORKBOOK = "./20240322-Electrical component library.xlsx"

# SCRIPT CONSTANTS

BOM_AUTHOR = "Justin Verkade"
TARGET_FILE = TARGET_CSV
TARGET_DIR = "/".join(TARGET_CSV.split("/")[:-1])
PROJECT_NAME = TARGET_FILE.split("/")[-1].split('.')[0]
BOM_FILE_FORMAT = "{}-{}_BOM.xlsx"
DATETIME_STRING = str(dt.datetime.now().date())
BOM_FILE_NAME = BOM_FILE_FORMAT.format(DATETIME_STRING, PROJECT_NAME)
OUTPUT_FILE_NAME = TARGET_DIR + "/" + BOM_FILE_NAME
print(OUTPUT_FILE_NAME)

# SCRIPT


def main():
    bill_of_materials = createTemplate()
    processKiCadBOM(bill_of_materials)
    bill_of_materials.save(OUTPUT_FILE_NAME)


def openLibrary(library_filename):
    library = pyxl.load_workbook(library_filename)
    sheet = library.active
    return library, sheet


def openDataframe(target_file):
    columns = ["Id",
               "Designator",
               "Footprint",
               "Quantity",
               "Designation",
               "Supplier and ref",
               "1",
               "2"]
    csv_dataframe = pd.read_csv(TARGET_FILE,
                                delimiter=';',
                                header=0,
                                names=columns)
    csv_dataframe = csv_dataframe.reset_index()
    csv_line_count = csv_dataframe.shape[0]
    return csv_dataframe, csv_line_count


def processKiCadBOM(bill_of_materials: pyxl.Workbook) -> pyxl.Workbook:
    sheet = bill_of_materials.active
    library, lib_sheet = openLibrary(LIBRARY_WORKBOOK)
    csv_dataframe, csv_line_count = openDataframe(TARGET_FILE)

    # log dataframe information
    print("File: %s.csv" % PROJECT_NAME)
    print("Components: %d\n" % csv_line_count)

    # log process status
    print("Start processing:")
    print("\rProgress [%60s][%3d%%]" % ("", 0), end="", flush=True)

    # process CSV
    actual_row = 0
    for index, row in csv_dataframe.iterrows():
        if "MountingHole" in row["Designation"]:
            continue
        if "Test" in row["Designation"]:
            continue

        # write references
        references = row["Designator"]
        sheet[f"C{actual_row + 8}"] = references
        sheet[f"C{actual_row + 8}"].alignment = Alignment(horizontal='left')

        # write quantities
        references = int(row["Quantity"])
        sheet[f"E{actual_row + 8}"] = references
        sheet[f"E{actual_row + 8}"].alignment = Alignment(horizontal='left')

        # processing flag
        processed = False

        # process resistors
        if row["Footprint"][:2] == "R_":
            fitting_component = None
            best_cost = 10.0e6

            # get target metrics
            footprint = row["Footprint"][2:6]
            resistance = ""
            tolerance = ""
            voltage = ""
            power = ""

            # for item in designation
            for val in row["Designation"].split(' '):
                # get voltage
                if val[-1] == "V":
                    voltage = val
                # get tolerance
                elif val[-1] == "%":
                    tolerance = val
                # get power value
                elif val[-1] == "W":
                    power = val
                else:
                    resistance = val

            # iterate all components
            library_row = 7
            while lib_sheet[f"B{library_row}"].value:
                library_type = str(lib_sheet[f"B{library_row}"].value)
                library_footprint = str(lib_sheet[f"E{library_row}"].value)
                library_resistance = str(lib_sheet[f"F{library_row}"].value)
                library_tolerance = str(lib_sheet[f"G{library_row}"].value)
                library_voltage = str(lib_sheet[f"H{library_row}"].value)
                library_power = str(lib_sheet[f"I{library_row}"].value)
                library_cost = float(lib_sheet[f"L{library_row}"].value)
                if library_type != "Resistor":
                    library_row += 1
                    continue
                if library_footprint != footprint:
                    library_row += 1
                    continue
                if resistance and library_resistance != resistance:
                    library_row += 1
                    continue
                if tolerance and library_tolerance != tolerance:
                    library_row += 1
                    continue
                if voltage and library_voltage != voltage:
                    library_row += 1
                    continue
                if power and library_power != power:
                    library_row += 1
                    continue
                if library_cost < best_cost:
                    fitting_component = library_row
                    best_cost = library_cost
                library_row += 1

            # write component information
            if fitting_component:
                # write manufacturer number
                manufacturer_number = lib_sheet[f"C{fitting_component}"].value
                alignment = Alignment(horizontal='left')
                sheet[f"A{actual_row + 8}"] = manufacturer_number
                sheet[f"A{actual_row + 8}"].alignment = alignment

                # write description
                description = "Resistor "
                description += footprint + " "
                description += resistance + " "
                description += voltage + " "
                description += tolerance + " "
                description += power
                alingment = Alignment(horizontal='left')
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = alingment

                # write costs
                costs = best_cost
                alignment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = alignment
                sheet[f"D{actual_row + 8}"].number_format = format

                # write total sum
                string = f"=D{actual_row + 8} * E{actual_row + 8}"
                alignment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"F{actual_row + 8}"] = string
                sheet[f"F{actual_row + 8}"].alignment = alignment
                sheet[f"F{actual_row + 8}"].number_format = format

                # write url
                url = lib_sheet[f"M{fitting_component}"].value
                alignment = Alignment(horizontal='left')
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = alignment
                sheet[f"G{actual_row + 8}"].style = "Hyperlink"
                processed = True

        # process capacitors
        elif row["Footprint"][:2] == "C_":
            fitting_component = None
            best_cost = 10.0e6

            # get target metrics
            footprint = row["Footprint"][2:6]
            capacitance = ""
            tolerance = ""
            voltage = ""

            # for item in designation
            for val in row["Designation"].split(' '):
                # get voltage
                if val[-1] == "V":
                    voltage = val
                # get tolerance
                elif val[-1] == "%":
                    tolerance = val
                else:
                    capacitance = val

            # iterate all components
            library_row = 7
            while lib_sheet[f"B{library_row}"].value:
                library_type = lib_sheet[f"B{library_row}"].value
                library_footprint = lib_sheet[f"E{library_row}"].value
                library_capacitance = lib_sheet[f"F{library_row}"].value
                library_tolerance = lib_sheet[f"G{library_row}"].value
                library_voltage = lib_sheet[f"H{library_row}"].value
                library_cost = float(lib_sheet[f"L{library_row}"].value)
                if library_type != "Capacitor":
                    library_row += 1
                    continue
                if library_footprint != footprint:
                    library_row += 1
                    continue
                if capacitance and library_capacitance != capacitance:
                    library_row += 1
                    continue
                if tolerance and library_tolerance != tolerance:
                    library_row += 1
                    continue
                if voltage and library_voltage != voltage:
                    library_row += 1
                    continue
                if library_cost < best_cost:
                    fitting_component = library_row
                    best_cost = library_cost
                library_row += 1

            # write component information
            if fitting_component:
                # write manufacturer number
                alignment = Alignment(horizontal='left')
                manufacturer_id = lib_sheet[f"C{fitting_component}"].value
                sheet[f"A{actual_row + 8}"] = manufacturer_id
                sheet[f"A{actual_row + 8}"].alignment = alignment

                # write description
                description = "Capacitor "
                description += footprint + " "
                description += capacitance + " "
                description += voltage + " "
                description += tolerance
                alignment = Alignment(horizontal='left')
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = alignment

                # write costs
                costs = best_cost
                alginment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = alginment
                sheet[f"D{actual_row + 8}"].number_format = format

                # write total sum
                sum_string = f"=D{actual_row + 8} * E{actual_row + 8}"
                alignment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"F{actual_row + 8}"] = sum_string
                sheet[f"F{actual_row + 8}"].alignment = alignment
                sheet[f"F{actual_row + 8}"].number_format = format

                # write url
                url = lib_sheet[f"M{fitting_component}"].value
                alignment = Alignment(horizontal='left')
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = alignment
                sheet[f"G{actual_row + 8}"].style = "Hyperlink"
                processed = True

        # process general
        else:
            # iterate all components
            library_row = 7
            while lib_sheet[f"B{library_row}"].value:
                # check if type is simular
                if lib_sheet[f"C{library_row}"].value != row["Designation"]:
                    library_row += 1
                    continue

                # write manufacturer number
                manufacturer_number = lib_sheet[f"C{library_row}"].value
                alignment = Alignment(horizontal='left')
                sheet[f"A{actual_row + 8}"] = manufacturer_number
                sheet[f"A{actual_row + 8}"].alignment = alignment

                # write description
                description = str(lib_sheet[f"B{library_row}"].value) + " "
                if sheet[f"E{library_row}"].value:
                    format = str(lib_sheet[f"E{library_row}"].value)
                    description += format + " "
                if sheet[f"F{library_row}"].value:
                    format = str(lib_sheet[f"F{library_row}"].value)
                    description += format + " "
                alignment = Alignment(horizontal='left')
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = alignment

                # write costs
                costs = float(lib_sheet[f"L{library_row}"].value)
                alignment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = alignment
                sheet[f"D{actual_row + 8}"].number_format = format

                # write total sum
                string = f"=D{actual_row + 8} * E{actual_row + 8}"
                alignment = Alignment(horizontal='left')
                format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                sheet[f"F{actual_row + 8}"] = string
                sheet[f"F{actual_row + 8}"].alignment = alignment
                sheet[f"F{actual_row + 8}"].number_format = format

                # write url
                url = lib_sheet[f"M{library_row}"].value
                alignment = Alignment(horizontal='left')
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = alignment
                sheet[f"G{actual_row + 8}"].style = "Hyperlink"
                processed = True
                break

        # fill with default if not in list
        if not processed:
            # write manufacturer number
            manufacturer_number = row["Footprint"] + ' - ' + row["Designation"]
            alignment = Alignment(horizontal='left')
            sheet[f"B{actual_row + 8}"] = manufacturer_number
            sheet[f"B{actual_row + 8}"].alignment = alignment

            # write costs
            costs = 0
            alignment = Alignment(horizontal='left')
            format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
            sheet[f"D{actual_row + 8}"] = costs
            sheet[f"D{actual_row + 8}"].alignment = alignment
            sheet[f"D{actual_row + 8}"].number_format = format

            # write total sum
            string = f"=D{actual_row + 8} * E{actual_row + 8}"
            alignment = Alignment(horizontal='left')
            format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
            sheet[f"F{actual_row + 8}"] = string
            sheet[f"F{actual_row + 8}"].alignment = alignment
            sheet[f"F{actual_row + 8}"].number_format = format

            # log failure
            print("\rFailed: %s%s" % (manufacturer_number, " " * 80))

        # increment row
        actual_row += 1

        # update progress bar
        percentage = int(100 * (index + 1) / csv_line_count)
        bar = "#" * int(60 * percentage / 100)
        print("\rProgress [%-60s][%3d%%]" % (bar, percentage),
              end="",
              flush=True)

    # log update finished
    percentage = 100
    bar = "#" * int(60 * percentage / 100)
    print("\rProgress [%-60s][%3d%%]" % (bar, percentage), end="", flush=True)
    print("\nProcessing completed!")
    print("─" * 80)
    print("Goodbye...")

    # write total cost formula
    alignment = Alignment(horizontal='left')
    sheet[f"E{actual_row + 8}"].value = "Total:"
    sheet[f"E{actual_row + 8}"].alignment = alignment

    alignment = Alignment(horizontal='left')
    format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
    sheet[f"F{actual_row + 8}"].value = f"=SUM(F8:F{actual_row + 7})"
    sheet[f"F{actual_row + 8}"].alignment = alignment
    sheet[f"F{actual_row + 8}"].number_format = format

    # close library document
    library.close()


def createTemplate() -> pyxl.Workbook:
    # log progress
    print("─" * 80)
    print("Create workbook ", end="", flush=True)

    # create workbook
    bill_of_materials = pyxl.Workbook()

    # log progress
    print("DONE")
    print("Create worksheet ", end="", flush=True)

    # create data sheet
    sheet = bill_of_materials.active
    sheet.title = PROJECT_NAME

    # write header
    sheet["A2"] = "Document:"
    sheet["A2"].alignment = Alignment(horizontal='left')
    sheet["A3"] = "Last changed:"
    sheet["A3"].alignment = Alignment(horizontal='left')
    sheet["A4"] = "Author:"
    sheet["A4"].alignment = Alignment(horizontal='left')
    sheet["B2"] = PROJECT_NAME
    sheet["B2"].alignment = Alignment(horizontal='left')
    sheet["B3"] = dt.datetime.now().date()
    sheet["B3"].alignment = Alignment(horizontal='left')
    sheet["B4"] = BOM_AUTHOR
    sheet["B4"].alignment = Alignment(horizontal='left')

    # write component column headers
    sheet["A6"] = "Components"
    sheet["A6"].alignment = Alignment(horizontal='left')
    sheet["A7"] = "Manufacturer number"
    sheet["A7"].alignment = Alignment(horizontal='left')
    sheet["B7"] = "Description"
    sheet["B7"].alignment = Alignment(horizontal='left')
    sheet["C7"] = "References"
    sheet["C7"].alignment = Alignment(horizontal='left')
    sheet["D7"] = "Costs"
    sheet["D7"].alignment = Alignment(horizontal='left')
    sheet["E7"] = "Qty"
    sheet["E7"].alignment = Alignment(horizontal='left')
    sheet["F7"] = "Total"
    sheet["F7"].alignment = Alignment(horizontal='left')
    sheet["G7"] = "Url"
    sheet["G7"].alignment = Alignment(horizontal='left')

    # set column widths
    sheet.column_dimensions['A'].width = 40     # Manufacturer number
    sheet.column_dimensions['B'].width = 50     # Description
    sheet.column_dimensions['C'].width = 40     # References
    sheet.column_dimensions['D'].width = 12     # Costs
    sheet.column_dimensions['E'].width = 12     # Qty
    sheet.column_dimensions['F'].width = 12     # Total
    sheet.column_dimensions['G'].width = 100    # Url

    # log progress
    print("DONE")
    print("─" * 80)

    return bill_of_materials


if __name__ == "__main__":
    main()
