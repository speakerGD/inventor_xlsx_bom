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
    "custom_length": ["Mass", "Масса_кг"],
}


def main():
    # Validate an inventor specification
    # Analyze the raw BOM
    # Calculate the resulting mass for all of the materials
    # Analyze the BOM template
    # Delete unused materials from the BOM template
    # Fill the BOM template

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

    # Process bill of materials
    if valid_for_materials(sys.argv[1]):
        bill_of_materials(sys.argv[1], sys.argv[2])

    # Process bill of purchases parts
    if valid_for_purchased(sys.argv[1]):
        bill_of_purchased(sys.argv[1], sys.argv[2])

    # Process bill of md1000 parts
    if valid_for_md1000(sys.argv[1]):
        bill_of_md1000(sys.argv[1], sys.argv[2])


def valid_for_materials(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to calculate BOM.
    Columns required:
    - bom_structure
    - quantity
    - material
    - mass
    """
    wb = openpyxl.load_workbook(source)

    raise NotImplementedError


def valid_for_purchased(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of purchased parts.
    Columns required:
    - part_number
    - bom_structure
    - quantity
    """
    raise NotImplementedError


def valid_for_md1000(source):
    """
    Check the inventor .xlsx specification `source` for necessary columns to derive a list of md1000 parts.
    Columns required:
    - part_number
    - description
    - quantity
    """
    raise NotImplementedError


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
