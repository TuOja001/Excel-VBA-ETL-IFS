# Automated Sales Data Processing with Excel and VBA
## Project Overview
This Excel project automates the process of handling sales data through data imports, consolidation, pivot tables, and VBA macros. It integrates data from multiple sources, processes it, and generates organized reports efficiently. 
Below is a detailed description of each function:

### Data Sources
The project consolidates data from the following sources into a single Excel file for processing:
- tuotteet.accdb: Product information from a Microsoft Access database
- ryhmat.xlsx: Group data stored in an Excel workbook
- myynnit.xlsx: Main sales dataset with transaction records
- myyjat.txt: Text document containing salesperson information

These sources are cleaned and formatted for analysis and reporting

### Tools and Functions
- Data Consolidation with XLOOKUP
- Pivot Tables for Summarization
- IFS Functions for Conditional Reporting
- Automated Report Generation with VBA Macro
- Text Functions for Data Cleaning and Formatting

### Data cleaning/Preparation
- Data Merging: The XLOOKUP function is used to link sales data from the myynnit sheet with related information from other sheets, ensuring accurate connections for seamless reporting.

- Pivot Tables: Pivot tables are created on the rapo worksheet to summarize key metrics like sales amounts and transaction counts, categorized by salesperson and group, providing clear and adjustable summaries.

- Conditional Summaries: IFS functions are utilized to generate conditional summary tables on the rapo sheet, calculating totals and transaction counts based on predefined conditions for each salesperson and group.

- Automation with VBA: A VBA macro (can be found in the file module3) named laske automates the report generation process, iterating through sales data to calculate totals and transaction counts. The results are inserted into the rapo sheet, summarizing each salesperson’s performance efficiently.

- Text Data Formatting: Various Excel text functions, such as Find, Left, Right, Proper, and Trim, are applied to clean and standardize text data, including generating email addresses in a specific format and ensuring consistent name capitalization. This preparation aids in accurate reporting and display.

### Exploratory Data Analysis
- What insights does the analysis reveal about sales performance, trends, and data quality?
- What patterns emerge that indicate which salespersons and groups consistently perform well?
- How do temporal trends provide opportunities for strategic planning and resource allocation?

### Data Analysis
- Top Performer: Rauno has the highest overall total
- Key Category: Äänentoisto is the leading category
- Balanced Contributions: Jenni and Sanni contribute evenly across categories
- Specialized Strengths: Hessu and Jani excel in specific categories such as Digiboxit
- Lower Performers: Sohvi and Valtteri have lower totals
- Growth Areas: Tarvikkeet and TV have lower overall totals

### Proposed action
Rauno's strong performance suggests that investing in his strengths could be beneficial, while the high activity in the Äänentoisto category indicates it should be prioritized for resources. Conversely, categories like Tarvikkeet and TV show lower totals but offer growth opportunities that require targeted investments.
