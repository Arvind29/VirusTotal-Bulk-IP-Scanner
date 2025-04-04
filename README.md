# VirusTotal-Bulk-IP-Scanner

Automates bulk IP reputation checks via VirusTotal API, outputting results with risk indicators in Excel.

Steps to Use the Script:

Prepare the Input File
  Create an Excel file (input.xlsx) in the same folder as the script.
  Ensure the file has a column named "IP" in column A with valid IP addresses.
The script will:
  1. Read input.xlsx.
  2. Fetch IP details from VirusTotal API.
  3. Save results in output.xlsx.
  4. Highlight IPs with Malicious > 2 in red.
  5. If an error occurs, it will display a message.
  6. Open output.xlsx in the same folder.
