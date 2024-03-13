# Matchinator

Matchinator is a Python tool designed to match products from a rejection list to their corresponding codes in an orderable list. This can be particularly useful in scenarios such as inventory management or data reconciliation where products need to be matched across different datasets.

# Features

Data Import: Matchinator can import data from Excel files for both rejection and orderable lists.
String Matching: It uses fuzzy string matching to find close matches between product names from the rejection list and product names from the orderable list.
Output Generation: Matchinator generates a new Excel workbook containing matched product names and their corresponding codes.

# Usage

- Ensure your rejection list is in an Excel file named rejets.xlsx with the sheet named rejets.
- Ensure your orderable list is in an Excel file named cadencier.xlsx with the sheet named Feuil1.
- After execution, a new Excel file named match.xlsx will be generated containing the matched products and their corresponding codes.

# Dependencies

- openpyxl: For reading and writing Excel files.
- difflib: For fuzzy string matching to find close matches.