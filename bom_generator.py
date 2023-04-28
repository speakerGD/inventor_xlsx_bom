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
    bill_of_materials(wb_source, wb_template)

    # Process bill of purchases parts
    bill_of_purchased(wb_source, wb_template)

    # Process bill of md1000 parts
    bill_of_md1000(wb_source, wb_template)


def bill_of_materials(source, template):
    """
    Check the inventor .xlsx specification `source` for necessary columns to calculate BOM. Columns required:
    - bom_structure
    - quantity
    - material
    - mass
    Calculate materials from the `source` and copy data to the template.
    Delete unused rows of materials from the template.
    """

    required_columns_ru = [
        "Структура спецификации",
        "КОЛ.",
        "Материал",
        "Масса",
    ]

    if not all_required_columns(source, required_columns_ru):
        print("Could not issue a bill of materials.")
        print("Source file must contain at least these columns:")
        for column in required_columns_ru:
            print(column)


def bill_of_purchased(source, template):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of purchased parts. Columns required:
    - part_number
    - bom_structure
    - quantity
    Copy purchased parts from the `source` to the template.
    """

    required_columns_ru = [
        "Обозначение",
        "Структура спецификации",
        "КОЛ.",
    ]

    if not all_required_columns(source, required_columns_ru):
        print("Could not issue a bill of purchased parts.")
        print("Source file must contain at least these columns:")
        for column in required_columns_ru:
            print(column)


def bill_of_md1000(source, template):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of md1000 parts. Columns required:
    - part_number
    - description
    - quantity
    Copy md1000 parts from the `source` to the template.
    """

    required_columns_ru = [
        "Обозначение",
        "Наименование",
        "КОЛ.",
    ]

    if not all_required_columns(source, required_columns_ru):
        print("Could not issue a bill of md1000 parts.")
        print("Source file must contain these columns:")
        for column in required_columns_ru:
            print(column)


def all_required_columns(source, columns):
    """
    Validate that all `columns` exist in the `source` xlsx file.
    """

    # Open the first sheet from the source workbook
    sheet = source[source.sheetnames[0]]

    for column in columns:
        exists = False

        if column in list(sheet.rows)[1]:
            exists = True

        # If at least one required column doesn't exist
        if not exists:
            return False

    return True


if __name__ == "__main__":
    main()
