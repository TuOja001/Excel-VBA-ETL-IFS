# Automated Sales Data Processing with Excel and VBA
Project Overview
This Excel project automates the process of handling sales data through data imports, consolidation, pivot tables, and VBA macros. 
It integrates data from multiple sources, processes it, and generates organized reports efficiently. 
Below is a detailed description of each function:

Key Features
Data Import and Organization: The project imports data from several sources, including Excel files, a text file, and an Access database. 
The data is placed in separate worksheets: myynnit for sales data, ryhmat for group data, myyjat for salespersons, and tuotteet for product information. 
Each dataset is carefully structured to facilitate further analysis and reporting.

Data Consolidation with XLOOKUP: The XLOOKUP function is used to merge data from different sheets. 
It links the sales data in the myynnit sheet with corresponding information from other sheets, such as salesperson and group data. 
This ensures that all the relevant details are accurately connected, enabling seamless reporting and analysis.

Pivot Tables for Summarization: Pivot tables are created on the rapo worksheet using the consolidated sales data from the myynnit sheet. 
These tables summarize key metrics, such as sales amounts and transaction counts, categorized by salesperson and group.
The pivot tables provide clear, concise summaries that can be adjusted for different reporting perspectives.

IFS Functions for Conditional Reporting: IFS functions are used to create conditional summary tables on the rapo sheet. 
These formulas calculate specific values based on pre-defined conditions, such as total sales and the number of transactions for each salesperson and group. 
The tables are formatted to ensure readability and accurate data representation.

Automated Report Generation with VBA Macro: A VBA macro, laske, automates the entire report generation process. 
The macro iterates through each salesperson and sales group in the myynnit worksheet, calculating totals and transaction counts. 
It uses two custom functions: yht, which calculates the total sales for a specific salesperson and group, and lkm, which counts the number of transactions. 
The results are inserted into the rapo sheet, where each salesperson’s performance is summarized. 
The macro is designed to handle large datasets efficiently, generating reports in a matter of seconds. 
A button on the worksheet allows users to trigger the macro easily.

Text Functions for Data Cleaning and Formatting: Various Excel text functions, such as Find, Left, Right, Proper, and Trim, are employed to clean and format the text data. 
Email addresses are automatically generated in the format firstname.lastname@utu.fi, all in lowercase. 
Names are formatted so that only the first letter is capitalized, ensuring consistency across the dataset. 
The first and last names are extracted from combined name fields for further data organization. 
These text manipulations standardize the data, preparing it for accurate reporting and display.
