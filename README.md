# Excel Email Extractor

A simple Python script that extracts email addresses from an Excel spreadsheet and saves them to a text file. This tool is particularly useful for extracting email addresses from Google Form responses exported to Excel.

## Features

- Extracts email addresses from a specified Excel column
- Works with Google Form response spreadsheets by default
- Saves extracted emails to a text file
- Handles empty cells gracefully
- Prints extracted emails to console

## Prerequisites

- Python 3.x
- openpyxl library

## Installation

1. Clone the repository:
```bash
git clone https://github.com/aatikah/email-extractor.git
cd email-extractor
```

2. Install the required package:
```bash
pip install openpyxl
```

## Usage

1. Place your Excel file in the same directory as the script
2. Update the `file_path` variable with your Excel file name:
```python
file_path = 'your_file_name.xlsx'
```

3. Run the script:
```bash
python email_extractor.py
```

The extracted emails will be:
- Printed to the console
- Saved to `emails.txt` in the same directory

## Configuration

You can customize the following parameters:
- `sheet_name`: Default is 'Form Responses 1' (standard Google Forms sheet name)
- `email_column_name`: Default is 'Email Address'
- `output_file`: Default is 'emails.txt'

## Example

```python
# Custom configuration
emails = extract_email_column(
    file_path='responses.xlsx',
    sheet_name='Sheet1',
    email_column_name='User Email'
)
```

## Error Handling

The script will raise a ValueError if:
- The specified email column is not found
- The Excel file cannot be opened
- The specified sheet name doesn't exist


## Contributing

Feel free to open issues or submit pull requests if you have suggestions for improvements.
---

**⭐ If you find this project helpful, please give it a star!**
- Support me- [buymeacoffee.com/aatikah](https://buymeacoffee.com/aatikah)
- Connect with me on LinkedIn: [LinkedIn Profile](https://www.linkedin.com/in/abdulhakeem-sulaiman/)
