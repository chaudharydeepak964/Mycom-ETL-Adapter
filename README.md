# Mycom CSV Parser

This repository contains a script to parse and transform raw CSV files into the required format for Mycom adapter integration. The code ensures data consistency, proper mapping, and compliance with Mycom’s schema before ingestion.

🔧 Features

1. Opens a CSV file for writing with UTF-8 encoding.
2. Writes the header row (column names) without quotes.
3. Loops through each row of the dataframe u70 using .itertuples().
4. For each value in the row:
5. If not empty/NaN → wrap it in double quotes "value".
6. If empty/NaN → replace with "N/A".
7. Writes rows into the CSV with Windows-style newlines (\r\n).

👉 Final output: a clean CSV where headers are unquoted, and all row values are quoted (or "N/A" if missing).