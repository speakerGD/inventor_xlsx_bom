import openpyxl
import sys
import os
from pathlib import Path


TITLES = {
    "part_number": ["Part Number", "Обозначение"],
    "bom_structure": ["BOM Structure", "Структура спецификации"],
    "quantity": ["QTY", "КОЛ."],
    "description": ["Description", "Наименование"],
    "material": ["Material", "Материал"],
    "mass": ["Mass", "Масса"],
    "stock_number": ["Stock Number", "Инвентарный номер"],
    "project": ["Project", "Проект"],
    "vendor": ["Vendor", "Поставщик"],
    "file_path": ["File Path", "Путь файла"],
    "custom_mass": ["Mass", "Масса_кг"],
}


def main():
    # Check that files are provided
    if (
        len(sys.argv) != 3
        or not sys.argv[1].endswith(".xlsx")
        or not sys.argv[2].endswith(".xlsx")
    ):
        sys.exit(
            "Usage: bom_generator.py <inventor_specification>.xlsx <bom_template>.xlsx"
        )

    # Check that provided inventor specification file exists
    if not os.path.exists("./" + sys.argv[1]):
        sys.exit(f"{sys.argv[1]} must be in the same folder with bom_generator.py")

    # Check that provided template file exists
    if not os.path.exists("./" + sys.argv[2]):
        sys.exit(f"{sys.argv[2]} must be in the same folder with bom_generator.py")

    # Load workbooks
    wb_source = openpyxl.load_workbook(sys.argv[1])
    wb_template = openpyxl.load_workbook(sys.argv[2])

    # Process bill of materials
    if valid_for_materials(wb_source):
        bill_of_materials(wb_source, wb_template)

    # Process bill of purchases parts
    if valid_for_purchased(wb_source):
        bill_of_purchased(wb_source, wb_template)

    # Process bill of md1000 parts
    if valid_for_md1000(wb_source):
        bill_of_md1000(wb_source, wb_template)


def valid_for_materials(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to calculate BOM.
    Columns required:
    - bom_structure
    - quantity
    - material
    - mass
    """

    # List of lists of possible names of required columns
    required_columns = []

    for title in TITLES:
        if title in ["bom_structure", "quantity", "material", "mass"]:
            required_columns.append(TITLES[title])

    return all_required_columns(source, required_columns)


def valid_for_purchased(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of purchased parts.
    Columns required:
    - part_number
    - bom_structure
    - quantity
    """

    # List of lists of possible names of required columns
    required_columns = []

    for title in TITLES:
        if title in ["part_number", "bom_structure", "quantity"]:
            required_columns.append(TITLES[title])

    return all_required_columns(source, required_columns)


def valid_for_md1000(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of md1000 parts.
    Columns required:
    - part_number
    - description
    - quantity
    """

    # List of lists of possible names of required columns
    required_columns = []

    for title in TITLES:
        if title in ["part_number", "description", "quantity"]:
            required_columns.append(TITLES[title])

    return all_required_columns(source, required_columns)


def all_required_columns(source, columns):
    """
    Validate that all `columns` exist in the `source` xlsx file.
    """
    # Open the first sheet from the source workbook
    sheet = source[source.sheetnames[0]]

    for column in columns:
        exists = False

        # For each possible name for the required column
        for column_name in column:
            if column_name in list(sheet.rows)[1]:
                exists = True

        # If at least one required column doesn't exist
        if not exists:
            return False

    return True


def bill_of_materials(source, template):
    """
    Calculate materials from the `source` and copy data to the template.
    Delete unused rows of materials from the template.
    """
    raise NotImplementedError


def bill_of_purchased(source, template):
    """
    Copy purchased parts from the `source` to the template.
    """
    raise NotImplementedError


def bill_of_md1000(source, template):
    """
    Copy md1000 parts from the `source` to the template.
    """
    raise NotImplementedError


if __name__ == "__main__":
    main()
