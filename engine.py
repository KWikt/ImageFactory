import copy
import os
import openpyxl
import xml.etree.ElementTree as eTree
import subprocess


# Start full process of making card from xlsx and svg template.
def process_starter(svg, xlsx, files_name, files_path, files_format, dpi):

    workbook = openpyxl.load_workbook(xlsx)
    sheet = workbook.active

    tree = eTree.parse(svg)
    root = tree.getroot()

    # Started from 3 because numbers (program) export xlsx with table name as row and second row is for column name.
    # Iterate over rows each loop give another row with data to process.
    for i in range(3, sheet.max_row + 1):
        new_root = root_modifier(sheet, root, i)

        if "%VAR_" in files_name:
            unique_name = name_receiver(sheet, i)
        else:
            unique_name = files_name

        if unique_name == "None":
            return
        else:
            if files_format == "svg":
                svg_maker(files_path, unique_name, new_root)
            elif files_format == "png":
                png_maker(files_path, unique_name, new_root, dpi)
            elif files_format == "pdf":
                pdf_maker(files_path, unique_name, new_root)


# Take xlsx data file as sheet, make copy of provided svg root and take row_number to iterate over specific row in xls.
# Return modified root.
def root_modifier(sheet, root, row_number):
    new_root = copy.deepcopy(root)
    for row in sheet.iter_rows(min_row=row_number, max_row=row_number):
        # Loop over row that have name with data to search.
        for column in sheet.iter_rows(min_row=2, max_row=2):
            # Take column name and current row value and save it then send to replace.
            for cell, name in zip(row, column):
                cell_value = str(cell.value)
                # Each text in svg should start with "%VAR_" and with "%".
                text_replace = "%VAR_" + str(name.value) + "%"
                picture_replace = str(name.value) + ".png"
                # Replace values method.
                replace_values(new_root, text_replace, picture_replace, cell_value)
    return new_root


# Replace values in xml_root with values from xlsx.
def replace_values(xml_root, name_to_change, new_picture_name, cell_value):
    # Inkscape delete from picture path some information that need to be recovered manually.
    # IF YOU(I) CHANGE FOLDER WHERE PICTURE ARE YOU NEED FIND WHAT INKSCAPE DELETE AND CHANGE THAT.
    namespace = {'sodipodi': 'http://sodipodi.sourceforge.net/DTD/sodipodi-0.dtd'}
    attribute = './/*[@sodipodi:absref]'
# Change text %VAR_some_text% to normal text in xls.
    for element in xml_root.iter():
        if element.text == name_to_change in element.text:
            element.text = element.text.replace(name_to_change, cell_value)
        # Search for image, take old path and make new with new file name from new_picture_name.
    for image_elem in xml_root.findall(attribute, namespace):
        old_full_path = image_elem.get("{http://sodipodi.sourceforge.net/DTD/sodipodi-0.dtd}absref")
        folder_path = os.path.dirname(old_full_path)
        new_image_path = os.path.join(folder_path, new_picture_name)
        if old_full_path == new_image_path:
            cleared_path = new_image_path.replace(new_picture_name, '')
            path_with_file = os.path.join(str(cleared_path), str(cell_value) + ".png")
            image_elem.set('{http://www.w3.org/1999/xlink}href', path_with_file)


# Return unique card name from column named "Name".
def name_receiver(sheet, row_number):
    for row in sheet.iter_rows(min_row=row_number, max_row=row_number):
        # Loop over row that have name with data to search.
        for column in sheet.iter_rows(min_row=2, max_row=2):
            # Take column name and current row value and save it.
            for cell, name in zip(row, column):
                card_name = str(cell.value)
                return card_name


# Check if chosen file name exists if not add number incremented for each new copy.
# Return name that doesn't exist.
def valid_name(directory, file_name, extension):
    increment = 1
    original_name = file_name
    # Iterate through files in the directory
    for root, dirs, files in os.walk(directory):
        for existing_file in files:
            if file_name + extension in files:
                if existing_file.startswith(original_name):
                    file_name = original_name + str(increment)
                    increment += 1
            else:
                # If the loop completes without finding the file
                return file_name + extension


# Take xml tree, name and save new file svg to provided path.
def svg_maker(file_path, file_name, xml_root):
    extension = ".svg"
    # Check names if exists add number before.
    file_name = valid_name(file_path, file_name, extension)

    file_path = os.path.join(file_path, file_name)

    xml_string = eTree.tostring(xml_root, encoding="utf-8", method="xml").decode()
    # Save the SVG string to the specified file.
    with open(file_path, 'w') as svg_file:
        svg_file.write(xml_string)


# Create png based on temporary SVG file from xml_root, change dpi for better resolution.
def png_maker(file_path, file_name, xml_root, dpi):
    extension = ".png"

    # Check names if exists add number before.
    file_name = valid_name(file_path, file_name, extension)
    png_file_path = os.path.join(file_path, file_name)

    # Create temporary SVG file.
    svg_file_path = os.path.join(file_path, "temp.svg")
    with open(svg_file_path, 'w') as svg_file:
        svg_file.write(eTree.tostring(xml_root).decode())

    # Create PDF file using Inkscape
    inkscape_cmd = [
        '/Applications/Inkscape.app/Contents/MacOS/inkscape',
        '--export-type=png',  # Change the export type to PNG
        '--export-filename={}'.format(png_file_path),
        '--export-dpi={}'.format(dpi),
        svg_file_path
    ]

    subprocess.run(inkscape_cmd)

    # Remove the temporary SVG file.
    os.remove(svg_file_path)


# Create pdf based on temporary SVG file from xml_root.
def pdf_maker(file_path, file_name, xml_root):
    extension = ".pdf"

    # Check names if exists add number before.
    file_name = valid_name(file_path, file_name, extension)
    pdf_file_path = os.path.join(file_path, file_name)

    # Create temporary SVG file.
    svg_file_path = os.path.join(file_path, "temp.svg")
    with open(svg_file_path, 'w') as svg_file:
        svg_file.write(eTree.tostring(xml_root).decode())

    # Create PDF file using Inkscape
    inkscape_cmd = [
        '/Applications/Inkscape.app/Contents/MacOS/inkscape',
        '--export-type=pdf',
        '--export-filename={}'.format(pdf_file_path),
        svg_file_path
    ]

    subprocess.run(inkscape_cmd)

    # Remove the temporary SVG file.
    os.remove(svg_file_path)


if __name__ == '__main__':

    provided_dpi = 500
    file_format = "png"
    name_file = "%VAR_Name%"
    out_path = "/Users/wiktorkrzywdzinski/Desktop/"
    svg_path = "/Users/wiktorkrzywdzinski/Desktop/CardMaker/SVG_Templates/Cards/SpellTemplate.svg"
    xls_path = "/Users/wiktorkrzywdzinski/Desktop/test.xlsx"

    process_starter(svg_path, xls_path, name_file, out_path, file_format, provided_dpi)
