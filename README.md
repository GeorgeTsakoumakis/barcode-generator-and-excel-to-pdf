# Barcode Generator for Excel and export to PDF
This app generates a barcode from given data, adds it to an excel at a specified cell, and exports to pdf

## How to run
1. Install python 3.X (latest version)
2. Clone this repository
3. Install the required packages using `pip install -r requirements.txt` in virtual environment

Windows
```
python -m venv .env
.env/Scripts/Activate.ps1
python -m pip install -r requirements.txt
python -m pip install -e .
```


Linux - I think?
```
python -m venv .env
source .env/Scripts/Activate
python -m pip install -r requirements.txt
python -m pip install -e .
```

## Important notes
1. Make sure to make a copy of the excel file in the root of the project
2. Change the filename and path where asked to
3. Change the cell where asked to
4. Change the data where asked to
5. Change the pdf filename where asked to
6. Create a venv before installing the required packages