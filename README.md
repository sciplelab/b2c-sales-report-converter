# B2C Sales Report Converter

This project is a Python Flask-based website where users can upload .csv files from Shopify to clean and map relevant columns to a .xlsx template file.

## Setup
1. Create a virtual environment using `python -m venv .venv`.
2. Activate it using `.\.venv\Scripts\Activate`.
3. Install the required packages using `pip install -r requirements.txt`.
4. Run the application using `python app.py`.

## Commands
Delete files in exports and imports folder
```sh
python app.py --clear
```
