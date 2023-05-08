# Initial properties for inventor_xlsx_bom project to work with.

PART_NUMBER = "Обозначение"
BOM_STRUCTURE = "Структура спецификации"
QUANTITY = "КОЛ."
DESCRIPTION = "Наименование"
MATERIAL = "Материал"
MASS = "Масса"
STOCK_NUMBER = "Инвентарный номер"
PROJECT = "Проект"
VENDOR = "Поставщик"
FILE_PATH = "Путь файла"
CUSTOM_MASS = "Mass"
CUSTOM_LENGTH = "Length"
CUSTOM_WIDTH = "Width"
CUSTOM_AREA = "Area"

PROFILE_MATERIAL = "ПрофильныйПрокат"
FLAT_MATERIAL = "ПлоскийПрокат"
PURCHASED = "ПокупныеИзделия"
MD1000 = "УнифицированныеИзделия"

FLAT_MATERIAL_PREFIX = (
    "Лист",
    "Пластина",
    "Плита",
    "Профнастил",
    "Фанера",
    "Сетка",
    "Сэндвич-панель",
)

# Keys must be in the template's B column's rows.
# Values are substrings of materials names that should be
# assosiated with those rows in the template.
MATERIAL_TYPE = {
    "Ст3сп": ["Ст3"],
    "08Х18Н10": ["08Х18Н10", "нерж"],
    "ЛС59-1": ["ЛС59", "латунь"],
    "Ф-4": ["Ф-4", "фторопласт"],
    "маслонаполненный ПА-6": ["капролон"],
    "08пс": ["08пс"],
}
