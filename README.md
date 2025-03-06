# Task Summary Project

This project contains a Google Apps Script that creates a summary sheet in a Google Spreadsheet. The script consolidates data from multiple sheets into a single summary sheet titled **MARCH SUMMARY**. It processes the data by reading from each sheet (ignoring the header rows) and calculates totals for each member across dates.

> **Credit:** The core script is credited to **Denis Mbote**.

## How It Works

1. **Script Overview:**
   - The script searches for all sheets except the one named **MARCH SUMMARY**.
   - It then reads data starting from row 3 (thus skipping constant title rows like "Region", "Team", "Member 1", "Member 2", "Total").
   - For each sheet, it sums the values from the **Total** column (column E) for two members (columns C and D).
   - It then creates a new summary sheet (or clears an existing one) titled **MARCH SUMMARY**.
   - The summary displays each unique member in the first column and the respective totals for each date in adjacent columns.

2. **User Interface:**
   - A custom menu titled **"Summarize Now"** is added to the Google Sheets UI.
   - When the user clicks **"Run Summary"** from the menu, the `createTaskSummary` function is executed.

## Deployment Instructions

1. **Open Your Google Spreadsheet:**
   - Open the Google Spreadsheet where you want to use this script.

2. **Access the Script Editor:**
   - Go to **Extensions > Apps Script**.

3. **Add the Code:**
   - Create a new script file (e.g., `Code.gs`) and paste the code from the `Code.gs` file in this repository.

4. **Save and Authorize:**
   - Save the project and authorize the script if prompted.

5. **Run the Script:**
   - Reload your spreadsheet.
   - You will see a new menu called **"Summarize Now"**. Click on **"Summarize Now" > "Run Summary"** to generate the summary.

## Credits

- **Denis Mbote** – Original script author and inspiration for this project.

---

Happy Summarizing!
