import argparse
import csv
import translator
import time


def main():
    # Parse arguments
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="ChatGPT Translator for Excels",
        epilog=f"See README.md for more details. Github: https://github.com/TaihouKai/chatgpt-excel-translator"
    )
    parser.add_argument(
        "-f",
        "--file",
        type=str,
        help="(REQUIRED) The excel file to read from. This file must be a valid excel file.",
    )
    parser.add_argument(
        "-m",
        "--model",
        type=str,
        default="gpt-3.5-turbo",
        help="(OPTIONAL) The model to use for translation (gpt-3.5-turbo, gpt-4). Default: gpt-3.5-turbo",
    )
    parser.add_argument(
        "-c",
        "--column",
        type=str,
        default="A",
        help="(OPTIONAL) The column of the excel file to read from. Default: A",
    )
    parser.add_argument(
        "-r",
        "--row",
        type=int,
        default=1,
        help="(OPTIONAL) The row of the excel file to start reading from. Default: 1",
    )
    parser.add_argument(
        "-o",
        "--out",
        type=str,
        default="translated",
        help="(OPTIONAL) Filename of the output file. Default: translated",
    )
    parser.add_argument(
        "-s",
        "--src",
        type=str,
        default="Japanese",
        help="(OPTIONAL) Source language. Default: Japanese",
    )
    parser.add_argument(
        "-d",
        "--dest",
        type=str,
        default="English",
        help="(OPTIONAL) Destination language. Default: English",
    )
    args = parser.parse_args()

    # Check if "file" param is set
    if args.file is None or args.file == "":
        print("Please specify a file to translate.")
        print("")
        parser.print_help()
        return

    # Check if file is xls or xlsx
    if not args.file.endswith(".xls") and not args.file.endswith(".xlsx"):
        print("Please specify a valid excel file.")
        print("")
        parser.print_help()
        return

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
        translated_text = translator.chatgpt_translate(text, args.model, from_lang=args.src, to_lang=args.dest)
        # Remove all newlines in translated_text
        translated_text = translated_text.replace("\n", " ").strip()
        translated.append(translated_text)
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
