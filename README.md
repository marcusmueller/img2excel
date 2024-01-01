# img2excel – convert images to excel spreadsheet art

## Usage

```
img2excel.py [-h] [-W WIDTH] [-H HEIGHT] INFILE [OUTFILE]

positional arguments:
  INFILE                Image file to read
  OUTFILE               Excel File to write (default: samename.xlsx)

options:
  -h, --help            show this help message and exit
  -W WIDTH, --width WIDTH
                        width of cell in pixels
  -H HEIGHT, --height HEIGHT
                        height of cell in pixels
```

## Dependencies

- PIL (for image loading)
- xlsxwriter (for excel file writing)
- Python > 3.6

## Example


