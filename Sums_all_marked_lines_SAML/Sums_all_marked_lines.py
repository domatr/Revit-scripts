# Date                = 09.MAY.2024
# Python              = 3.0-3.6
# Revit               = 2023
# IronPython          = 3.4.1 (3.4.1.1000)
# RevitPythonShell    = 2.0.2
# Name: Sums_all_marked_lines
import clr

# Add Reference to Revit API
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Function to get the length of a detail line in meters
def get_detail_line_length(line):
    # Revit's internal unit for length is feet, so convert to meters
    return line.GeometryCurve.Length * 0.3048

# Main execution
def main():
    # Set up the script environment
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    # Get the selected elements in Revit
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        print("No elements selected.")
        return

    # Dictionary to store lengths by line type
    line_lengths = {}
    total_length = 0

    for id in selected_ids:
        element = doc.GetElement(id)
        if isinstance(element, DetailLine):
            length = get_detail_line_length(element)
            line_type_name = element.LineStyle.Name

            # Add to the total length and length by line type
            total_length += length
            if line_type_name in line_lengths:
                line_lengths[line_type_name] += length
            else:
                line_lengths[line_type_name] = length

    # Print the lengths by line type and the total length
    for line_type, length in line_lengths.items():
        print(f"Total length for line type '{line_type}': {length:.2f} meters")
    print(f"Total length of all selected detail lines: {total_length:.2f} meters")

# Execute the main function
main()