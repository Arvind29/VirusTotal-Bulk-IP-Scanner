# VirusTotal-Bulk-IP-Scanner
Automates bulk IP reputation checks via VirusTotal API, outputting results with risk indicators in Excel.

Steps to Use the Script
1. Prepare the Input File
Create an Excel file (input.xlsx) in the same folder as the script.

Ensure the file has a column named "IP" in column A with valid IP addresses.

The script will:
  Read input.xlsx.
  Fetch IP details from VirusTotal API.
  Save results in output.xlsx.
  Highlight IPs with Malicious > 2 in red.
  If an error occurs, it will display a message.
  Check Output File
  Open output.xlsx in the same folder.
