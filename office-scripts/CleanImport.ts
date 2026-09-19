/**
 * CleanImport - Office Script (no install, no macros)
 *
 * Excel on the web: Automate tab > New Script > paste > Save.
 * Also runs from the Automate tab of desktop Excel on a Microsoft 365 licence.
 *
 * Use this one when company policy blocks VBA macros: Office Scripts are not
 * macros and are usually allowed where .xlsm files are not.
 *
 * Trims stray whitespace, strips the non-breaking spaces that silently break
 * VLOOKUP, and turns numbers-stored-as-text back into real numbers.
 */
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getUsedRange();
  if (!range) {
    console.log("Sheet is empty - nothing to clean.");
    return;
  }

  const values = range.getValues();
  const formulas = range.getFormulas();
  let textFixed = 0;
  let numFixed = 0;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      // never overwrite a formula with its own cached result
      if (String(formulas[r][c]).startsWith("=")) continue;

      const v = values[r][c];
      if (typeof v !== "string" || v.length === 0) continue;

      const cleaned = v
        .replace(/ /g, " ")        // non-breaking space
        .replace(/[‎‏]/g, "") // directional marks
        .replace(/\s+/g, " ")
        .trim();

      if (cleaned !== v) {
        values[r][c] = cleaned;
        textFixed++;
      }

      // "13 500 000" and "13500000" both become a real number
      const numeric = cleaned.replace(/\s/g, "").replace(",", ".");
      if (numeric.length > 0 && !isNaN(Number(numeric))) {
        values[r][c] = Number(numeric);
        numFixed++;
      }
    }
  }

  range.setValues(values);
  console.log(`Cleaned ${textFixed} text cell(s), converted ${numFixed} to numbers.`);
}
