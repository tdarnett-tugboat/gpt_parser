# GPT Parser

## Installation and Setup

This script requires Python 3.9 or later. If python is already installed, move to step 4.

You can download the latest version of Python from the official Python website: https://www.python.org/downloads/

Alternatively, Mac users can install python through Homebrew.

1. Install Homebrew if you don't already have it by opening Terminal and entering the following command:
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```
2. Once Homebrew is installed, you can use it to install Python 3.9 or newer by entering the following command:
```bash
brew install python@3.9
```

3. After the installation is complete, verify that Python 3.9 (or newer) is installed by entering the following command:

```bash
python3 --version
```

4. Create a new virtual environment for the project:
```bash
python3 -m venv myenv
```

5. Activate the virtual environment
```bash
source myenv/bin/activate
```

6. Install the required packages:

```bash
pip install -r requirements.txt
```

You're now set up and can run the script!

## Running the script
1. Ensure you're in the virtual environment
```bash
source myenv/bin/activate 
```

2. Call the script with the configuration options
```bash
python gpt_cell_replacer.py --input_file /path/to/input_file.xlsx --sheet Sheet1 --output_file /path/to/output_file.xlsx --api_key YOUR_OPENAI_API_KEY
```
Replace `/path/to/input_file.xlsx` and `/path/to/output_file.xlsx` with the actual file paths of your input and output spreadsheets, respectively. Also, replace `Sheet1` with the name of the sheet in your input file that you want to read, and replace `YOUR_OPENAI_API_KEY` with your actual OpenAI API key.

