# Remove Duplicate Entries Based on Email Column

## Description
This Excel script is designed to sort and remove duplicate entries from a table based on the email addresses listed in the first column. It sorts the table by the second column in descending order, then filters out duplicate email addresses, keeping only the first occurrence of each. Finally, it deletes all rows except for the unique ones and logs the changes. This script is useful for managing lists where unique email addresses are required.

## Script

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet
  let worksheet = workbook.getActiveWorksheet();

  // Get the first table on the sheet
  let table = worksheet.getTables()[0];

  // Sort the table by the second column in descending order
  table.getSort().apply([
    {
      key: 1, // The second column "Submission Time"
      ascending: false // Sort in descending order
    }
  ]);

  console.log("Sorted the table by the second column in descending order.");

  // Get the data range between the header and total row
  let dataRange = table.getRangeBetweenHeaderAndTotal();

  // Get the values from the data range
  let dataValues = dataRange.getValues();

  // Create an object to hold unique emails
  let uniqueEntries: { [email: string]: boolean } = {};

  // Filter the data for unique email entries
  let uniqueDataValues = dataValues.filter((row) => {
    let email = row[0] as string;
    if (uniqueEntries.hasOwnProperty(email)) {
      return false;
    } else {
      uniqueEntries[email] = true;
      return true; // Keep this email
    }
  });

  // Clear the existing table rows (except headers)
  const rowCount = table.getRowCount();
  if (rowCount > 1) { // Check if there's more than just the header row
    table.getRangeBetweenHeaderAndTotal().delete(ExcelScript.DeleteShiftDirection.up);
  }

  // Add the new rows to the table
  table.addRows(null, uniqueDataValues);

  console.log("Removed duplicate rows based on the email column.");
}
````

## Important Notes
- :warning: **DANGER** This script modifies the existing table, so make a copy of your data before running if you need to retain the original data.
- Ensure that the table has a header row and that the first column contains email addresses.
- The script sorts data based on the second column. If the "Submission Time" is located in a different column, adjust `key: 1` to match the appropriate column index (note that column indices start from 0).
