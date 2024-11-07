# P-L-statement-in-excel
Profit &amp; Loss (P&amp;L) Sheet Creation from Scratch

This guide covers how to create a P&L sheet from raw data and consolidate it into a single database for analysis. The process involves data cleaning, formatting, unique identifier creation, and multi-year data integration.

Steps for P&L Sheet Creation

**Step 1: Understanding and Structuring the Data**
The data consists of three fiscal years—2020, 2021, and 2022—each containing five columns:
P&L Account: Type of account (e.g., Net Sales, Direct Costs)
Parent Company: Predefined parent company details
Name of Parent Company**: Actual name of the parent company
Amount: Financial figures related to each P&L account
Account Number: Unique identifier for each account

**Step 2: Data Formatting and Cleaning**
**Naming and Aligning Columns**: Ensure columns are consistently named and aligned for easy viewing. Use the shortcut Alt + H + A + L for left alignment.
**Adding Filters**: To apply filters, select the column headers and press Alt + A + T.
**Removing Totals Rows**: Filter out the "Total" rows in the "Name of Parent Company" column to avoid duplicate values. Once filtered, select all the "Total" rows and press Ctrl + - to delete them.
**Custom Formatting**: Add borders around cells and bold the column headers for readability.

**Step 3: Creating Unique Identifiers**
To handle duplicate Account Numbers across different companies, create a unique code: Add a new column called Code.
Use concatenation to combine No. of Parent Company and Account Number. Formula options: **=CONCAT(F4, C4) || =F4&C4** This generates a unique identifier for each P&L entry.

**Creating the Consolidated Database**

**Add a New Sheet**: Create a sheet named Database.
**Copy Unique Codes**: Copy the unique Code column from each fiscal year sheet (2020, 2021, 2022) and paste them into the Database sheet.
**Removing Duplicates**: Select the Code column, go to the Data tab, then select Data Tools > Remove Duplicates.
**Add Database Columns**: Structure the database with these columns:
Code, P&L Account, No. of Partner Company, Name of Partner Company, FY2020, FY2021, FY2022

**Use Formulas to Populate Data**

P&L Account:
=VLOOKUP($A2, 'FY2020'!$B$4:$G$60, 2, FALSE)
=VLOOKUP($A58, 'FY2021'!$A$3:$F$85, 2, FALSE)
=VLOOKUP($A91, 'FY2022'!$A$3:$F$77, 2, FALSE)

No. of Partner Company:
=VLOOKUP($A2, 'FY2020'!$B$4:$G$60, 3, FALSE)
=VLOOKUP($A58, 'FY2021'!$A$3:$F$85, 3, FALSE)
=VLOOKUP($A91, 'FY2022'!$A$3:$F$77, 3, FALSE)

Name of Partner Company:
=VLOOKUP($A2, 'FY2020'!$B$4:$G$60, 4, FALSE)
=VLOOKUP($A59, 'FY2021'!$A$3:$F$85, 4, FALSE)
=VLOOKUP($A91, 'FY2022'!$A$3:$F$77, 4, FALSE)

Fiscal Year Amounts:
=-SUMIF('FY2020'!$A:$A, Database!$A2, 'FY2020'!$E:$E)
=-SUMIF('FY2021'!$A:$A, Database!$A2, 'FY2021'!$E:$E)
=-SUMIF('FY2022'!$A:$A, Database!$A2, 'FY2022'!$E:$E)

**Categorizing Data**
Add categories under P&L Account to group similar entries. Common categories include:
  Net Sales
  Other Revenues
  Direct Costs
  Personnel Costs
  Other Operating Expenses (OPEX)
  Depreciation & Amortization (D&A)
  Financial Items
  Extraordinary Items
  Taxes
  
By following these steps, you will have a well-organized and structured P&L sheet and consolidated database, allowing for efficient data analysis across fiscal years.

**P&L Statement Creation**
**Step 1: Creating the P&L Statement Format**
Define Categories and Calculations:
  Create a template with the following financial categories:
    Total Revenues
    Gross Margin
    EBITDA (Earnings Before Interest, Taxes, Depreciation, and Amortization)
    EBIT (Earnings Before Interest and Taxes)
    EBT (Earnings Before Taxes)
    Net Income
Arrange these categories to reflect the flow from revenue to net income.

**Step 2: Pulling Data Using SUMIF**
To populate each category with data from the Database sheet, use the SUMIF function to reference specific rows:
  FY2020: =SUMIF(Database!$I:$I, $B4, Database!E:E)
  FY2021: =SUMIF(Database!$I:$I, $B4, Database!F:F)
  FY2022: =SUMIF(Database!$I:$I, $B4, Database!G:G)
Here, $B4 should refer to the specific category (e.g., Gross Margin) being retrieved. Adjust the row and column references as needed to match your template layout.

**Step 3: Calculating Year-on-Year Percentage Variations**
  Variation % FY21-22: =IF(ISERROR((D4/C4) - 1), "n.a.", IF(((D4/C4) - 1) > 1, ">100.0%", IF(((D4/C4) - 1) < -1, "<-100.0%", (D4/C4) - 1)))
  Variation % FY22-22: =IF(ISERROR((E4/D4) - 1), "n.a.", IF(((E4/D4) - 1) > 1, ">100.0%", IF(((E4/D4) - 1) < -1, "<-100.0%", (E4/D4) - 1)))
This formula calculates the percentage change from one year to the next, displaying values as “n.a.” if an error occurs, ">100.0%" if the increase exceeds 100%, and "<-100.0%" if the decrease exceeds -100%.

**Step 4: Key Performance Indicators (KPIs)**
Add KPIs to measure profitability and operational efficiency:
  Gross Margin %: =Gross Margin / Total Revenue
  EBITDA %: =EBITDA / Total Revenue
  EBIT %: =EBIT / Total Revenue

**Step 5: Final Formatting**
  Apply consistent formatting to ensure readability (e.g., bold headers, currency formatting for financial values, percentage format for variations).
  Add cell borders and adjust column widths as needed.

**Step 6: Verifying Net Income with VLOOKUP**
To confirm the Net Income value, use a VLOOKUP formula to pull data directly from the Database sheet:
  =VLOOKUP("Net Income", Database!A:G, column_number, FALSE)
Replace column_number with the column index for the relevant fiscal year (e.g., FY2022).

By following these steps, you will create a comprehensive P&L statement that supports year-on-year analysis, highlights key financial metrics, and ensures data accuracy.
