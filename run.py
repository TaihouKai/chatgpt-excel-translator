import argparse
import csv
import translator


def main():
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="ChatGPT Translator for Excels",
        epilog="See README.md for more details."
    )
    parser.add_argument("-f", "--file", type=str, default="sample.xlsx", help="The excel file to read from")
    parser.add_argument("-c", "--column", type=str, default="F", help="The column of the excel file to read from")
    parser.add_argument("-r", "--row", type=int, default=3, help="The row of the excel file to start reading from")
    parser.add_argument("-o", "--out", type=str, default="translated", help="Filename of the output file")
    args = parser.parse_args()

    translated = []
    
    # Parse data
    try:
        data = translator.read_xlsx(args.file)
        data_parsed = translator.parse_xlsx(data, args.column, args.row)
    except Exception as e:
        print(f"{e}")
        print("")
        parser.print_help()
        return

    # Remove nan and None
    temp = []
    for data in data_parsed:
        data = str(data)
        if data == "nan":
            data = ""
        if data == "None":
            data = ""
        temp.append(data)
    data_parsed = temp
    
    for text in data_parsed:
        if text is None or str(text) == "nan" or str(text) == "None":
            text = ""
        translated.append(translator.chatgpt_translate(text))
    
    rows = zip(data_parsed, translated)

    with open(f"{args.out}.csv", 'w', newline='') as f:
        csv_writer = csv.writer(f)
        for row in rows:
            csv_writer.writerow(row)
    
    print("Done!")


if __name__ == "__main__":
    main()
