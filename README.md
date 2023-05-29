# Payment-Claim-Letter

This is a Python script that generates payment claim letters using data from a CSV file. It utilizes the pandas, docxtpl, os, and num2words libraries to read the data, populate a template, and convert numerical values to words.

Prerequisites
Before running the script, make sure you have the following installed:

Python (version 3.6 or higher)
pandas library
docxtpl library
num2words library
You can install the required libraries using the following command:

Copy code
pip install pandas docxtpl num2words
Usage
Prepare the Data

Create a CSV file named DebtDetail.csv with the following columns: 'Account', 'Name', 'Unit', 'Amount', 'Document', 'Class', 'Doc. date', 'Due date'. Each row represents a debt record.
Save the CSV file in the same folder as the script.
Prepare the Template

Create a Word document named PaymentClaimLetter.docx with a table.
The table should have the following columns: 'Document', 'Unit', 'Class', 'Doc. date', 'Due date', 'Amount'.
Add placeholders ({{Name}}, {{Business}}, {{Debt}}, {{DebtLetter}}) in the Word document where the corresponding data will be inserted.
Run the Script

Execute the script using the following command:
Copy code
python script.py
Output

The script will generate a payment claim letter for each unique account in the CSV file.
The output files will be saved in the same folder as the script with the naming format: Account-Name.docx.
Functionality
Reading Data

The script reads the data from the CSV file DebtDetail.csv using the pandas library.
It extracts the relevant columns 'Account', 'Name', 'Unit', and 'Amount'.
It calculates the total debt amount for each account.
Conversion of Debt Amount to Words

The function ConvertToWords converts the debt amount from numerical form to words using the num2words library.
It splits the amount into the integer and decimal parts.
The integer part is converted to words using the num2words function with the English language.
The decimal part is formatted as a fraction (e.g., "50/100") or as "zero" if it is zero.
Populating the Template

The script uses the docxtpl library to render the Word document template PaymentClaimLetter.docx.
It iterates over each unique account and populates the template with the corresponding data.
The template is populated with values for placeholders {{Name}}, {{Business}}, {{Debt}}, and {{DebtLetter}}.
Generating the Output

The populated template is saved as a new Word document using the account number and name as the filename.
The resulting documents will contain the payment claim letter with the relevant data and the debt amount converted to words.
Customization
You can customize the script according to your specific requirements:

Adjust the CSV file structure and column names to match your data.
Modify the template (PaymentClaimLetter.docx) to match your desired letter format and content.
Change the language for converting the debt amount to words by modifying the lang parameter in the ConvertToWords function.
Author
Leandro DB overrun115@gmail.com
License
This project is licensed under the MIT License.
