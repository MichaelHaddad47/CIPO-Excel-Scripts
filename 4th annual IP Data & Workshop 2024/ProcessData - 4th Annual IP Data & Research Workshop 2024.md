# Dietary Restrictions and Hearing Processing

## Script

```typescript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Get the table on the active worksheet
  let table = workbook.getActiveWorksheet().getTables()[0];

  // Get the data body range of the table (excluding headers and totals)
  let dataBodyRange = table.getRangeBetweenHeaderAndTotal();

  // Get the values of the entire table
  let tableValues = dataBodyRange.getValues();

  // Loop through each row of the table
  for (let i = 0; i < tableValues.length; i++) {
    // Process Dietary Restrictions
    let restrictions = tableValues[i][6] ? tableValues[i][6].toString() : ""; // Ensure we're working with a string
    processDietaryRestrictions(restrictions, dataBodyRange, i);

    // Process Hearing
    let hearing = tableValues[i][14] ? tableValues[i][14].toString() : ""; // Ensure we're working with a string
    processHearing(hearing, dataBodyRange, i);
  }

  console.log("Processed both Dietary Restrictions and Hearing columns.");
}

function processDietaryRestrictions(restrictionsString: string, dataBodyRange: ExcelScript.Range, rowIndex: number) {
  // Normalize the string: remove brackets, split by ',', and trim each element
  let parsedRestrictions = restrictionsString
    .replace(/^\["/, '')  // Remove leading '["'
    .replace(/"\]$/, '')  // Remove trailing '"]'
    .split('","')         // Split by '","'
    .map(s => s.trim())   // Trim whitespace from each element
    .map(s => s.replace(/^ \t/, '')); // Remove any leading " \t" from each entry

  let dairyValue = parsedRestrictions.some(r => r.includes("Dairy") || r.includes("Produits laitiers")) ? "✔" : "";
  let peanutsValue = parsedRestrictions.some(r => r.includes("Peanuts") || r.includes("Cacahuètes")) ? "✔" : "";
  let eggsValue = parsedRestrictions.some(r => r.includes("Eggs") || r.includes("Œufs")) ? "✔" : "";
  let glutenValue = parsedRestrictions.some(r => r.includes("Gluten")) ? "✔" : "";
  let treeNutsValue = parsedRestrictions.some(r => r.includes("Tree nuts") || r.includes("Fruits à coque")) ? "✔" : "";
  let soyProductsValue = parsedRestrictions.some(r => r.includes("Soy products") || r.includes("Produits à base de soja")) ? "✔" : "";

  // For 'Other', find any entry that doesn't match known categories
  let otherEntries = parsedRestrictions.filter(r =>
    !r.includes("Dairy") && !r.includes("Produits laitiers") &&
    !r.includes("Peanuts") && !r.includes("Cacahuètes") &&
    !r.includes("Eggs") && !r.includes("Œufs") &&
    !r.includes("Gluten") &&
    !r.includes("Tree nuts") && !r.includes("Fruits à coque") &&
    !r.includes("Soy products") && !r.includes("Produits à base de soja")
  ).join(", ");
  let otherValue = otherEntries ? otherEntries : "";

  dataBodyRange.getCell(rowIndex, 7).setValues([[dairyValue]]);
  dataBodyRange.getCell(rowIndex, 8).setValues([[peanutsValue]]);
  dataBodyRange.getCell(rowIndex, 9).setValues([[eggsValue]]);
  dataBodyRange.getCell(rowIndex, 10).setValues([[glutenValue]]);
  dataBodyRange.getCell(rowIndex, 11).setValues([[treeNutsValue]]);
  dataBodyRange.getCell(rowIndex, 12).setValues([[soyProductsValue]]);
  dataBodyRange.getCell(rowIndex, 13).setValues([[otherValue]]);
}


function processHearing(hearingString: string, dataBodyRange: ExcelScript.Range, rowIndex: number) {
  if (hearingString.startsWith('[') && hearingString.endsWith(']')) {
    let cleanedHearingString = hearingString.replace(/'/g, '"');
    let parsedResponses: string[] = JSON.parse(cleanedHearingString);

    let ciponetValue = parsedResponses.includes("CIPOnet") || parsedResponses.includes("OPICnet") ? "✔" : "";
    let cipoinfoValue = parsedResponses.includes("CIPOinfo") || parsedResponses.includes("OPICinfo") ? "✔" : "";
    let colleagueValue = parsedResponses.includes("Colleague") ? "✔" : "";
    let managerValue = parsedResponses.includes("Manager") ? "✔" : "";
    let directorValue = parsedResponses.includes("Director") ? "✔" : "";
    let interConnexValue = parsedResponses.includes("InterConnex") ? "✔" : "";
    let cipoconnexValue = parsedResponses.includes("CIPOConnex") || parsedResponses.includes("OPICConnex") ? "✔" : "";

    // For 'Other2', filter out known categories and join remaining entries
    let otherEntries = parsedResponses.filter(r =>
      !r.includes("CIPOnet") && !r.includes("OPICnet") &&
      !r.includes("CIPOinfo") && !r.includes("OPICinfo") &&
      !r.includes("InterConnex") &&
      !r.includes("Colleague") &&
      !r.includes("Manager") &&
      !r.includes("Director") &&
      !r.includes("CIPOConnex") && !r.includes("OPICConnex")
    ).join(", ");
    let other2Value = otherEntries ? otherEntries : "";

    dataBodyRange.getCell(rowIndex, 15).setValues([[ciponetValue]]);
    dataBodyRange.getCell(rowIndex, 16).setValues([[cipoinfoValue]]);
    dataBodyRange.getCell(rowIndex, 17).setValues([[interConnexValue]]);
    dataBodyRange.getCell(rowIndex, 18).setValues([[cipoconnexValue]]);
    dataBodyRange.getCell(rowIndex, 19).setValues([[colleagueValue]]);
    dataBodyRange.getCell(rowIndex, 20).setValues([[managerValue]]);
    dataBodyRange.getCell(rowIndex, 21).setValues([[directorValue]]);
    dataBodyRange.getCell(rowIndex, 22).setValues([[other2Value]]);
  } else {
    console.log(`Row ${rowIndex + 1}: Invalid hearing string format`);
  }
}
