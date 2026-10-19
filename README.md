# ImageOrganizer

A small `openpyxl` script for a Shopify-oriented product spreadsheet. It takes extra image values from numbered columns and places them in image rows beneath each product, writing a new workbook instead of replacing the input.

## Expected workbook

- Save the source file as `input.xlsx` beside `script.py`.
- Use a worksheet named `Sheet1` with a header row.
- Include an `Images` column for each product's first image.
- Put additional images in columns with numeric headers such as `2`, `3`, and `4`.
- Keep the product handle two columns before `Images`; the script references that column when arranging rows.

The script assumes this layout and does not offer command-line options. Review a copy of your workbook before using the output in another system.

## Run

```sh
python -m pip install openpyxl
python script.py
```

The output is `input_updated.xlsx` in the same directory. The script adds an `Image Index` column and may insert rows to make space for extra images. It leaves `input.xlsx` in place.

**Known limitation:** In a basic workbook with one product and two extra images, the inserted image rows had blank `Handle` cells. Check and fill the generated rows as required by your destination import format. The script does not import anything into Shopify.

No license file is currently included.
