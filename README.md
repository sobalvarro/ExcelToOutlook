# ExcelToOutlook
Excel tool to create mass e-mail with custom fields from a table

Description: The tool reads an Outlook template and replaces the "words" in the columns with the values in the table, then displays all the generated messages for a final revision.

The tool also adds the "To:" e-mail field, "CC:", Subject and add a file as attachment if a valid path is provided. The tool generates a unique e-mail for every line in the table. The "customfields" in blank will be ignored 

Important: Do not change the order of the columns and the rows. If you need more fields to be retrieved you can add columns. The columns needs to be named uniquely

Instructions:  

1) Open Outlook and create a message, substitute the words or phrases that you want to be unique for each e-mail with the column headers in the table. Please check carefully that the words you want to be substituted match exactly with the column header text (an extra space for example could lead to errors).  
2) Save the outlook message as a Outlook Template (*.oft)  
3) Click the "Select Template" Button in the spreadsheet and select the outlook template.  
4) Click the "Generate Emails" Button.  
5) Review generated e-mails and click send manually for each of them  


If an error occurred check that all the columns are named correctly, that there are no columns with empty rows or columns and that the headers are on the row 10
