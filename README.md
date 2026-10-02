# Address Excel Converter

A small Windows desktop app that combines one address per text-file line with four shared shipping fields and writes an `.xlsx` workbook. It handles 1,000 rows in one pass and runs locally; the only Python package is `openpyxl`.

## Start

1. Install Python 3.10 or newer from [python.org](https://www.python.org/downloads/windows/). During setup, enable the Python Launcher (`py`).
2. Double-click `run.bat`. On first launch it creates a local `.venv` and installs `openpyxl`; an internet connection is needed for that initial install.
3. Choose the address text file and fixed-fields CSV, select where to save the workbook, and click **Process and create Excel**.

## Input formats

The text file must have one address per line, with these fields in order:

```text
Name, Street, City, State, Zip
```

Comma, tab, semicolon, and pipe delimiters are supported. A header row (`Name,Street,City,State,Zip`) is optional. For comma-delimited rows with extra commas, the first field is the name, the last three are city/state/ZIP, and the fields in between are joined as the street. Quoted CSV fields are supported. Un-delimited or incomplete lines are reported and prevent export, rather than being guessed. ZIP codes are kept as text so leading zeroes are preserved.

The fixed-fields CSV contains exactly one data row with four nonblank values, in this order:

```csv
Service,Package_Type,Billing,Package_Weight
Ground,Parcel,Prepaid,1
```

The header is optional. The same four values are copied to every output row. Replace or choose another CSV whenever those values change.

## Output

The workbook has these columns in the required order:

```text
Customer_Name | AddressOne | Country | Postal | City | State | Service | Package_Type | Billing | Package_Weight
```

Country is set to `US`. The sheet includes a frozen header row and filters. The converter does not upload or transmit any data.

## Optional command line

To run without the desktop window, install the requirement and call the functions in `app.py` from your own Python script. `parse_addresses`, `read_fixed_fields`, and `write_workbook` are exposed for that use.