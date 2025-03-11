# monday-data-export
## Description
A tool to help Monday.com users export data from Monday.com API.

## Supported output formats
* CSV
* Excel
* Google Sheets

## Installation
1. Prepare monday-data-export.conf by referencing monday-data-export.conf.example
2. Prepare monday_api_token.txt by placing your Monday API token
3. Prepare google_credentials.json by referencing google_credentials.json.example (if you want to output to Google Sheets. if not, just skip this step.)
4. pip install -r requirements.txt (to check and prepare necessary python modules)
5. run monday-data-export.py

## Data filtering
Monday.com exports data that includes learning source links and group names. This tool helps filter out those elements from the output, ensuring that the final dataset contains only one table header and the intended data.
To exclude specific group names, add them to the exclusion list in monday-data-export.conf.

## Contact
Any questions please drop me an email or follow me at X.


