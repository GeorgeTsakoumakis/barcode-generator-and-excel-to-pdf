# Barcode Generator for Excel and export to PDF
This app generates a barcode from given data, adds it to an excel at a specified cell, and exports to pdf

## How to run
1. Install python 3.X (latest version)
2. Clone this repository
    * To clone this repository from the command line, cd into your directory and run `git clone` + https link of repository found at the green `Code` button
3. Create a virtual environment and install the required packages
    * To create a virtual environment, run `python -m venv .env` in the root of the project
    * To activate the virtual environment, run `.env/Scripts/Activate.ps1` in the root of the project
    * To install the required packages, run `python -m pip install -r requirements.txt` in the root of the project

Windows
```
python -m venv .env
.env/Scripts/Activate.ps1
python -m pip install -r requirements.txt
```


Linux - I think?
```
python -m venv .env
source .env/Scripts/Activate
python -m pip install -r requirements.txt
```

4. Install the required packages using `pip install -r requirements.txt` in virtual environment

## How to run
You can simply run the app by running `main.py` in the root of the project. From the command line, run `python main.py`. You can also run the app from an IDE like PyCharm or VSCode but the latter requires the Python extention.

## Important notes
1. Make sure to make a copy of the excel file in the root of the project
2. Change the filename and path where asked to
3. Change the cell where asked to
4. Change the data where asked to
5. Change the pdf filename where asked to
6. Create a venv before installing the required packages