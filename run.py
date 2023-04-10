import argparse
import csv
import translator
import time


def main():
    # Parse arguments
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="ChatGPT Translator for Excels",
        epilog="See README.md for more details.",
    )
    parser.add_argument(
        "-f",
        "--file",
        type=str,
        default="sample.xlsx",
        help="The excel file to read from",
    )
    parser.add_argument(
        "-c",
        "--column",
        type=str,
        default="A",
        help="The column of the excel file to read from",
    )
    parser.add_argument(
        "-r",
        "--row",
        type=int,
        default=1,
        help="The row of the excel file to start reading from",
    )
    parser.add_argument(
        "-o",
        "--out",
        type=str,
        default="translated",
        help="Filename of the output file",
    )
    args = parser.parse_args()

    translated = []

    # Parse data
    try:
        print(f"Reading file: {args.file} ...")
        data = translator.read_xlsx(args.file)
        data_parsed = translator.parse_xlsx(data, args.column, args.row)
    except Exception as e:
        print("")
        print(f"{e}")
        print("")
        parser.print_help()
        return
    # Replace nan and None
    temp = []
    for data in data_parsed:
        data = str(data)
        if data == "nan":
            data = ""
        if data == "None":
            data = ""
        temp.append(data)
    data_parsed = temp

    print("")

    print(f"Entries to be translated: {len(data_parsed)}")

    # Translate
    count = 0
    sleep_counter = 0
    sleep_max = 500
    for text in data_parsed:
        # Sleep every 500 requests
        if sleep_counter == sleep_max:
            print(f"Sleeping for 10 seconds ...")
            time.sleep(10)
            sleep_counter = 0
        count += 1
        # Although the data is parsed, still check for nan and None anyway
        if text is None or str(text) == "nan" or str(text) == "None":
            text = ""
        print(f"Translating: {count} / {len(data_parsed)} ...")
        # Default: Japanese -> English
        translated.append(translator.chatgpt_translate(text))
        sleep_counter += 1

    print("")

    print(f"Translation finished. Writing to file: {args.out}.csv ...")

    # Transpose data
    rows = zip(data_parsed, translated)
    # Write data to csv
    with open(f"{args.out}.csv", "w", newline="") as f:
        csv_writer = csv.writer(f)
        for row in rows:
            csv_writer.writerow(row)

    print("Done.")


if __name__ == "__main__":
    main()
