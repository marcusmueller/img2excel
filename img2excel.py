#!/usr/bin/env python3
# SPDX-License-Identifier: EUPL-1.2
# Copyright 2024 Marcus Müller

import xlsxwriter
import argparse
from PIL import Image


def main(
    infile: Image.Image, outfile: str, column_width: int = 20, row_height: int = 20
) -> None:
    width: int = infile.width
    height: int = infile.height
    image = infile.convert("RGB")

    colors = set(image.getdata())

    workbook = xlsxwriter.Workbook(outfile)
    color_map = {
        color: workbook.add_format(
            {"bg_color": f"#{color[0]:02x}{color[1]:02x}{color[2]:02x}"}
        )
        for color in colors
    }

    sheet = workbook.add_worksheet()
    sheet.set_zoom(100 * max(0.1, min(4, 1920 / (width * column_width))))
    sheet.set_column_pixels(0, width, column_width)
    for row in range(height):
        sheet.set_row_pixels(row, row_height)
        for col in range(width):
            color: tuple = image.getpixel((col, row))
            sheet.write(row, col, "", color_map[color])

    workbook.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser("image to excel")
    parser.add_argument(
        "INFILE", type=argparse.FileType("rb"), help="Image file to read"
    )
    parser.add_argument(
        "OUTFILE",
        nargs="?",
        help="Excel File to write (default: samename.xlsx)",
        default=None,
    )
    parser.add_argument(
        "-W", "--width", type=int, help="width of cell in pixels", default=20
    )
    parser.add_argument(
        "-H", "--height", type=int, help="height of cell in pixels", default=20
    )
    args = parser.parse_args()
    infile = Image.open(args.INFILE)
    if not args.OUTFILE:
        outfile = args.INFILE.name + ".xlsx"
    else:
        outfile = args.OUTFILE
    main(infile, outfile, args.width, args.height)
