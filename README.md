# video-coding-parsing
quick python to extract video coding data from separate spreadsheets


## Setup

```
python3 -m venv venv
source ./venv/bin/activate
pip install -r requirements.txt
```

## Usage

```
logsheets_to_tabular.py [xlsx files...]
```

Provide paths to files to process as arguments.  CSV data with one row per
sheet is provided on stdout.  CSV data with input and data warnings is
provided on stderr.

## Example

```
logsheets_to_tabular.py input_xlsx_files/*.xlsx > observations.csv 2> warnings.csv
```
