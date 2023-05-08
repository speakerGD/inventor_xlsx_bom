# Inventor-XLSX-BOM

#### Video Demo: https://youtu.be/S6HBZ5QyhrQ

#### Description:

This program is designed for specific usage in the company I work for as a design engineer. A part of my workflow is to issue a bill of materials, a bill of purchased parts and a bill of company's "md1000" parts for each products I design. These bills must be issued as one xlsx file in accordance with the particular template.

Each Inventor Assembly FIle (.iam) contains a specification (Bill of Materials), there listed all the parts presented in the product. There are a lot of different data that can be listed in that specification for each part, like part number, bom structure, quntity, discription and so on. It's also possible to add custom data using Inventor iLogic module. I use custom data to calculate amout of used materials for the product.

In my case I derive:

- mass of parts with up to three decimal points
- length of parts made of profile materials
- lenght, width and area of parts made of flat materials

Inventor specification can be exported as an xlsx file. That helps to parse it to derive information needed to issue the document mentioned above. Doing it manually is a tedious process, expecially than it comes to big complicated products with lots of data. This program semi-automates this process.

**Note:** _the program uses openpyxl module, which works only with Excel 2010 files._

**What the program does in core:**

1. Load a specification and a template.
   - Specification as a source workbook.
   - Template as a template workbook.
2. Proceed a bill of materials, a bill of purchased parts and a bill of md1000 parts. For each of the bills:
   - Filter rows of the specification in a particular way.
   - Retrieve data from related columns.
   - Summorize the data.
   - Fill the data into the template workbook's worksheet.
3. Save the resulting template workbook as a new xlsx file.

All important initial properties are encapsulated in settings.py. The properties like titles of columns in a specification and titles of worksheets in the template make the program flexible to handle possible changes in the template and specifications. Flat material prefixes and material type allows to scale the program to work with a wider range of materials.
