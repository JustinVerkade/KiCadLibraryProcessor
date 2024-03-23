from openpyxl.styles import Alignment, numbers
import openpyxl as pyxl
import tkinter as tk
import datetime as dt
import pandas as pd
    
############################
###### USER CONSTANTS ######
############################

TARGET_DIR          = ".\\"
TARGET_CSV          = "FoxConnect.csv"

###############################
###### LIBRARY CONSTANTS ######
###############################

LIBRARY_WORKBOOK    = ".\\20240322-Electrical component library.xlsx"

##############################
###### SCRIPT CONSTANTS ######
##############################

BOM_AUTHOR          = "Justin Verkade"
TARGET_FILE         = TARGET_DIR + "\\" + TARGET_CSV
PROJECT_NAME        = TARGET_CSV.split('.')[0]
BOM_FILE_FORMAT     = "{}-{} BOM.xlsx"
OUTPUT_FILE_NAME    = TARGET_DIR + "\\" + BOM_FILE_FORMAT.format(str(dt.datetime.now().date()), PROJECT_NAME)

####################
###### SCRIPT ######
####################

def main():
    bill_of_materials = createTemplate(TARGET_DIR, TARGET_CSV)
    bill_of_materials = processKiCadBOM(bill_of_materials)


    bill_of_materials.save(OUTPUT_FILE_NAME)

def processKiCadBOM(bill_of_materials:pyxl.Workbook) -> pyxl.Workbook:
    # get current sheet
    sheet = bill_of_materials.active

    # open library document
    library = pyxl.load_workbook(LIBRARY_WORKBOOK)
    library_sheet = library.active

    # load CSV dataframe
    columns = ["Id","Designator","Footprint","Quantity","Designation","Supplier and ref", "1", "2"]
    csv_dataframe = pd.read_csv(TARGET_FILE, delimiter=';', header=0, names=columns)
    csv_dataframe = csv_dataframe.reset_index()

    # process CSV
    actual_row = 0
    for index, row in csv_dataframe.iterrows():

        # check if mounting hole
        if "MountingHole" in row["Designation"]:
            continue

        # check if test point
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
            while library_sheet[f"B{library_row}"].value:
                # check type
                if str(library_sheet[f"B{library_row}"].value) != "Resistor":
                    library_row += 1
                    continue
                # check footprint
                if str(library_sheet[f"E{library_row}"].value) != footprint:
                    library_row += 1
                    continue
                # check capacitance
                if str(resistance and library_sheet[f"F{library_row}"].value) != resistance:
                    library_row += 1
                    continue
                # check tolerance
                if str(tolerance and library_sheet[f"G{library_row}"].value) != tolerance:
                    library_row += 1
                    continue
                # check voltage
                if str(voltage and library_sheet[f"H{library_row}"].value) != voltage:
                    library_row += 1
                    continue
                # check power
                if str(power and library_sheet[f"I{library_row}"].value) != power:
                    library_row += 1
                    continue
                # get lowest costfitting_component
                component_cost = float(library_sheet[f"L{library_row}"].value)
                if component_cost < best_cost:
                    fitting_component = library_row
                    best_cost = component_cost
                library_row += 1
            
            # write component information
            if fitting_component:
                # write manufacturer number
                manufacturer_number = library_sheet[f"C{fitting_component}"].value
                sheet[f"A{actual_row + 8}"] = manufacturer_number
                sheet[f"A{actual_row + 8}"].alignment = Alignment(horizontal='left')
                
                # write description
                description = "Resistor "
                description += footprint + " "
                description += resistance + " "
                description += voltage + " "
                description += tolerance + " "
                description += power
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = Alignment(horizontal='left')

                # write costs
                costs = best_cost
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"D{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write total sum
                sheet[f"F{actual_row + 8}"] = f"=D{actual_row + 8} * E{actual_row + 8}"
                sheet[f"F{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"F{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write url
                url = library_sheet[f"M{fitting_component}"].value
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = Alignment(horizontal='left')
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
            while library_sheet[f"B{library_row}"].value:
                # check type
                if library_sheet[f"B{library_row}"].value != "Capacitor":
                    library_row += 1
                    continue
                # check footprint
                if library_sheet[f"E{library_row}"].value != footprint:
                    library_row += 1
                    continue
                # check capacitance
                if capacitance and library_sheet[f"F{library_row}"].value != capacitance:
                    library_row += 1
                    continue
                # check tolerance
                if tolerance and library_sheet[f"G{library_row}"].value != tolerance:
                    library_row += 1
                    continue
                # check voltage
                if voltage and library_sheet[f"H{library_row}"].value != voltage:
                    library_row += 1
                    continue
                # get lowest costfitting_component
                component_cost = float(library_sheet[f"L{library_row}"].value)
                if component_cost < best_cost:
                    fitting_component = library_row
                    best_cost = component_cost
                library_row += 1
            
            # write component information
            if fitting_component:
                # write manufacturer number
                manufacturer_number = library_sheet[f"C{fitting_component}"].value
                sheet[f"A{actual_row + 8}"] = manufacturer_number
                sheet[f"A{actual_row + 8}"].alignment = Alignment(horizontal='left')
                
                # write description
                description = "Capacitor "
                description += footprint + " "
                description += capacitance + " "
                description += voltage + " "
                description += tolerance
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = Alignment(horizontal='left')

                # write costs
                costs = best_cost
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"D{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write total sum
                sheet[f"F{actual_row + 8}"] = f"=D{actual_row + 8} * E{actual_row + 8}"
                sheet[f"F{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"F{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write url
                url = library_sheet[f"M{fitting_component}"].value
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"G{actual_row + 8}"].style = "Hyperlink"
                processed = True

        # process general
        else:
            # iterate all components
            library_row = 7
            while library_sheet[f"B{library_row}"].value:
                # check if type is simular
                if library_sheet[f"C{library_row}"].value != row["Designation"]:
                    library_row += 1
                    continue

                # write manufacturer number
                manufacturer_number = library_sheet[f"C{library_row}"].value
                sheet[f"A{actual_row + 8}"] = manufacturer_number
                sheet[f"A{actual_row + 8}"].alignment = Alignment(horizontal='left')
                
                # write description
                description = str(library_sheet[f"B{library_row}"].value) + " "
                if library_sheet[f"E{library_row}"].value:
                    description += str(library_sheet[f"E{library_row}"].value) + " "
                if library_sheet[f"F{library_row}"].value:
                    description += str(library_sheet[f"F{library_row}"].value) + " "
                # if library_sheet[f"G{library_row}"].value:
                #     description += str(library_sheet[f"G{library_row}"].value) + " "
                # if library_sheet[f"H{library_row}"].value:
                #     description += str(library_sheet[f"H{library_row}"].value)
                sheet[f"B{actual_row + 8}"] = description
                sheet[f"B{actual_row + 8}"].alignment = Alignment(horizontal='left')

                # write costs
                costs = float(library_sheet[f"L{library_row}"].value)
                sheet[f"D{actual_row + 8}"] = costs
                sheet[f"D{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"D{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write total sum
                sheet[f"F{actual_row + 8}"] = f"=D{actual_row + 8} * E{actual_row + 8}"
                sheet[f"F{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"F{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
                
                # write url
                url = library_sheet[f"M{library_row}"].value
                sheet[f"G{actual_row + 8}"].hyperlink = url
                sheet[f"G{actual_row + 8}"].value = url
                sheet[f"G{actual_row + 8}"].alignment = Alignment(horizontal='left')
                sheet[f"G{actual_row + 8}"].style = "Hyperlink"
                processed = True
                break

        # fill with default if not in list
        if not processed:
            # write manufacturer number
            manufacturer_number = row["Footprint"] + ' - ' + row["Designation"]
            sheet[f"B{actual_row + 8}"] = manufacturer_number
            sheet[f"B{actual_row + 8}"].alignment = Alignment(horizontal='left')
                    
            # write costs
            costs = 0
            sheet[f"D{actual_row + 8}"] = costs
            sheet[f"D{actual_row + 8}"].alignment = Alignment(horizontal='left')
            sheet[f"D{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE
            
            # write total sum
            sheet[f"F{actual_row + 8}"] = f"=D{actual_row + 8} * E{actual_row + 8}"
            sheet[f"F{actual_row + 8}"].alignment = Alignment(horizontal='left')
            sheet[f"F{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE

        # increment row
        actual_row += 1

    # write total cost formula
    sheet[f"E{actual_row + 8}"].value = "Total:"
    sheet[f"E{actual_row + 8}"].alignment = Alignment(horizontal='left')
    sheet[f"F{actual_row + 8}"].value = f"=SUM(F8:F{actual_row + 7})"
    sheet[f"F{actual_row + 8}"].alignment = Alignment(horizontal='left')
    sheet[f"F{actual_row + 8}"].number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE

    # close library document
    library.close()
    return bill_of_materials

def createTemplate(target_dir:str, target_csv:str) -> pyxl.Workbook:
    # create workbook
    bill_of_materials = pyxl.Workbook()
    
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

    return bill_of_materials

if __name__ == "__main__":
    main()