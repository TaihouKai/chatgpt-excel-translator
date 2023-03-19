import openai
import pandas


def read_xlsx(file: str) -> pandas.DataFrame:
    """Read an excel file and return a pandas dataframe"""
    return pandas.read_excel(file, header=None, engine="openpyxl")


def parse_xlsx(data: pandas.DataFrame, column_excel: str, start_row_excel: int) -> list:
    """Parse the data from the pandas dataframe"""
    column_number = ord(column_excel) - 65
    start_row = start_row_excel - 1
    return data[column_number][start_row:].tolist()


def chatgpt_translate(text: str, from_lang: str = "Japanese", to_lang: str = "English") -> str:
    """Send a message to the chatbot and return the response"""
    ## Null line?
    if text is None or text == "nan" or text == "None":
        return ""

    ## Empty line?
    if text.strip() == "":
        return ""

    messages = [
        {"role": "system", "content": f"You are a translator. You only receive the {from_lang} text and you must translate it into {to_lang}."},
        {"role": "user", "content": f"{text}"}
    ]
    result = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )
    return result["choices"][0]["message"]["content"]

