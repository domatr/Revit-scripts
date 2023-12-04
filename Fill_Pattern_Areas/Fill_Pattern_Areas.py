# -*- coding: utf-8 -*-

#### INTRO 03

__title__           = "Analyzes all selected Fill Pattern and their area."
__doc__             = """Version = 0.1

Date                = 03.DEC.2023
Python              = 3.0-3.6
Revit               = 2023
IronPython          = 3.4.1 (3.4.1.1000)
RevitPythonShell    = 2.0.2

_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
Description:
Analyzes all selected Fill Patterns, calculates their respective areas, 
totals for each type, and the overall total from all selected Fill Patterns.
_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
How-to:
-> Select only the Fill Pattern you want.
-> Run the script.
-> The calculation results will be displayed in the Terminal.
-> A file named 'fill_pattern_areas.csv' will be saved to your desktop.
_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
Last update:
- [03.DEC.2023] - 0.1 RELEASE
_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
To-Do:
- Name and save as Excel file
_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
Author:     Doan, Manh Trung (DMT)
LICENSE:    CC BY-NC-SA"""

#### Code 34

import clr
import os
import csv
import datetime

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter, FilledRegion
from Autodesk.Revit.UI import TaskDialog
from collections import defaultdict
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

def get_selected_fill_pattern_areas():
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        TaskDialog.Show('Error', 'No elements selected.')
        return

    # Collecting Fill Pattern.
    filled_regions = FilteredElementCollector(doc, selected_ids).OfClass(FilledRegion)

    type_areas = defaultdict(float)
    total_area = 0.0
    output_data = []

    # Getting the area of each Fill Pattern
    for region in filled_regions:
        area_param = region.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        area = area_param.AsDouble() if area_param else 0
        area_m2 = round(area * 0.09290304, 2)   # Convert from sqf to sqm and round to two decimal places
        total_area += area_m2

        # Getting the type name
        region_type = doc.GetElement(region.GetTypeId())
        type_name = region_type.Name if region_type else "Unknown Type"
        type_areas[type_name] += area_m2
        output_data.append([region.Id.ToString(), type_name, f"{area_m2:.2f} m2"])

    # Display Terminal
    print("----Fill Pattern----\n")
    for type_name, area in type_areas.items():
        output_data.append(["Total for Type: " + type_name, f"{area:.2f} m2"])
    print("\n".join(', '.join(row) for row in output_data))
    print("\n")
    print(f"Total Area, {total_area:.2f} m2")

    # Get the current date and time
    current_datetime = datetime.datetime.now()
    formatted_datetime = current_datetime.strftime("%Y%m%d_%H%M%S")     # Date and Time format

    # Define the path for the CSV file
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    csv_file_name = f"Fill_Pattern_{formatted_datetime}.csv"             # Name the file
    csv_file_path = os.path.join(desktop_path, csv_file_name)

    # Write the data to the CSV file
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Element ID', 'Type Name', 'Area (m2)'])  # Write header
        for row in output_data:
            writer.writerow(row)

    print(f"CSV file saved to {csv_file_path}")

get_selected_fill_pattern_areas()