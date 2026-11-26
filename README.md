# Data Cleaning Sample: Titanic Passenger Dataset

A sample data-cleaning project on a public dataset (Titanic passenger list, 891 rows, from the Data Science Dojo public datasets repo: https://github.com/datasciencedojo/datasets). Built to show how I clean and document messy tabular data.

## What it does
- Checks and removes duplicates (none found)
- Trims and normalises text fields
- Splits `Name` into Surname, Title and Given names, groups rare titles
- Fills missing Age (177 rows) with the median by Title and Class, and flags those rows in `Age_Missing`
- Fills 2 missing Embarked values with the most common port
- Keeps Cabin as Deck / Has_Cabin (77% missing, not imputed)
- Adds readable labels, FamilySize, AgeGroup, and flags Fare outliers (not removed)
- Missing cells: 866 before, 0 after. Every step is logged in the `Cleaning_Log` sheet.

## Files
- `clean.py` - the full script (pandas)
- `titanic_raw.csv` - original data
- `titanic_clean.csv` - cleaned data
- `titanic_cleaned.xlsx` - workbook with Summary, Cleaning_Log, Missing_Before, Clean_Data and a raw sample

Run: `pip install pandas openpyxl && python clean.py`

Data source and license belong to their original authors. This is a practice sample, not client work.
