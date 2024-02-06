function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet
  let worksheet = workbook.getActiveWorksheet();

  // Get the first table on the sheet
  let table = worksheet.getTables()[0];

  // Sort the table by the second column in descending order
  table.getSort().apply([
    {
      key: 1, // The second column
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
    let email = row[0] as string; // Cast the email to a string to satisfy TypeScript's type checking
    if (uniqueEntries.hasOwnProperty(email)) {
      return false; // This email has already been encountered, skip it
    } else {
      uniqueEntries[email] = true; // Mark this email as encountered
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
