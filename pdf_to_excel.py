import argparse
import os
import tabula
import pandas as pd




def main():
    pdf_path, pages, output_path = parseArguments()

def parseArguments():
    parser = argparse.ArgumentParser("Extract tables from PDF and export to Excel for DCF modeling")

    parser.add_argument("pdf_path", type=str, help="Path to input PDF file eg. input/example.pdf")
    parser.add_argument("--pages", type=str, default="all", help="Pages to extract")

    args = parser.parse_args()

    pdf_path = args.pdf_path
    pages = args.pages

    name_no_ext = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = f"output/{name_no_ext}_tables.xlsx"
    
    print(f"Input PDF: {pdf_path}")
    print(f"Pages: {pages}")
    print(f"Output Excel: {output_path}")

    return pdf_path, pages, output_path


if __name__ == "__main__":
    main()