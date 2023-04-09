# Barcode Generator for Excel and export to PDF
This app generates a barcode from given data, adds it to an excel at a specified cell, and exports to pdf

***Version:*** 1.0.0

***Author:*** Georgios Tsakoumakis

***Contact Information:***
- email: gtsakoumakis2004@gmail.com
- school email: g.tsakoumakis2@newcastle.ac.uk

***Copyright:*** © Georgios Tsakoumakis 2023. All rights reserved. This app is licensed under the MIT License. You may use this app as long as you mention the author Georgios Tsakoumakis and provide a link to the repository. Any unauthorized use or distribution of this app or its source code is strictly prohibited.

This software can be used to develop other apps and is free to use for personal and commercial use. Only condition is that when using the contents of this repository the name of the author and link to the repository are provided and use of the source code is mentioned in the documentation. Also, the application should not be sold as is. If you want to sell the application, please contact me first.

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
5. Create a venv before installing the required packages