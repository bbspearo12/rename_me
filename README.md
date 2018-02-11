# Rename Me
This is a script that renames AWS billing statements from their original downloaded 
formats to human readable format with customer names and billing month and year.
For instance `6760-2473901168722-San Jose Water Company-312628085799-2017-12-statement` will be renamed to `San Jose Water Company AWS_Dec 2017`

## Flow

Script requires 3 parameters as command line arguments
* Zipfile consisting of billing statements
* Renaming template excel sheet (has to be of a particular format, see below)
* Output directory to which renamed files are to be copied

Script will extract the zip file to a directory with the same name and directory as the zipfile.
It will then get each file from the directory, read the account number and fetch the accountname from the file.
File will then be copied to the required output folder with the account name previously fetched.

**Note**

## Requirements

### Original Billing statement naming convention
```sh
* The original billing file should contain account number in the 4th token (when tokenzied by `-`)
```

### Template Excel sheet Requirements
```sh
Script assumes that the renaming template is a particular format
* There is a sheet named `Renaming Template` (case sensitive)
* Its 1st column should be headers
* Column A is account numbers and Column D is New name
```

### System Requirements
``` sh
* Python >= 2.7.10
* openpyxl v2.5.0 python module (pip install openpyxl==2.5.0) 
```
### Example Usage

#### Invocation
```sh
python rename_me.py --template billing-template.xlsx --outputdir renamed-statements --zipfile billing-artifacts.zip
```

#### output
```sh
Extracted billing statemets from billing-artifacts.zip to billing-artifacts
Account 098333046595 not found in template file, skipping 6600-2473901166150-PDF Solutions-098333046595-2017-12-statement.pdf
Account 044855270273 not found in template file, skipping 5500-2680059592816-Groupware Cloud Ops-044855270273-2017-12-statement.pdf
Account 421571602166 not found in template file, skipping 6008-2680059595330-Groupware IT-421571602166-2017-12-statement.pdf
Account 044613621107 not found in template file, skipping 7250-3023656978497-Dolby Laboratories, Inc.-044613621107-2017-12-statement.pdf
Account 062052777684 not found in template file, skipping 8038-3985729653288-LeanTaas-062052777684-2017-12-statement.pdf
Account 410844597293 not found in template file, skipping 5658-2680059593704-Ampush-410844597293-2017-12-statement.pdf
Account 004612015045 not found in template file, skipping 7250-3023656978506-Dolby Laboratories, Inc.-004612015045-2017-12-statement.pdf
Account 070866847466 not found in template file, skipping 5789-2680059594272-Augmedix-070866847466-2017-12-statement.pdf
Account 046531007673 not found in template file, skipping 7250-3023656981043-Dolby Laboratories, Inc.-046531007673-2017-12-statement.pdf

**** Summary ****
Billing statements zip file: billing-artifacts.zip
Renaming Template billing-template.xlsx
Target folder renamed-statements
Files to rename 57
Successfully renamed 48
Files for which account could not be found 9
**** End of Summary ****
```