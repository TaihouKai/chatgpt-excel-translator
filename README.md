# ChatGPT Translator for Excels

This script is a simple integration of using ChatGPT on Excel forms, to translate texts from columns.

This script is designed to help translators, as they often receive xlsx files as original text files.

# Setup

1. Login to OpenAI Dashboard: https://platform.openai.com/
2. Create an API Key: https://platform.openai.com/account/api-keys > "Create new secret key" > Copy the key
3. Create a new environmental variable based on your operating system, which name being `OPENAI_API_KEY`, and value being `<your_api_key>`
* * Windows: https://www.architectryan.com/2018/08/31/how-to-change-environment-variables-on-windows-10/
* * Linux: https://www.cyberciti.biz/faq/set-environment-variable-linux/
* * Mac: https://www.twilio.com/blog/2017/01/how-to-set-environment-variables.html
4. In terminal, enter this repository's directory and execute `pip install -r requirements.txt`

> If you'd like to modify default values of the parameters (i.e., arguments), please modify `parser.add_argument(..., default=xxx, ...)` in `run.py` according to your own need.
> Or, you can just specify the parameters manually.

# Run

> Remove square brackets from real commands. They simply mean "inputs".

1. Put Excel file to be translated in the root directory of this repository.
2. `python run.py -f [file_to_translate.xlsx]`

# Other Parameters

Run `python run.py -h` to read the manual.
