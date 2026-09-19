/**
 * AgingReport - Office Script (no install, no macros)
 *
 * Excel on the web: Automate tab > New Script > paste > Save.
 *
 * Reads a table with a due-date column and an outstanding-amount column, then
 * writes a receivables aging summary (not due / 0-30 / 31-60 / 61-90 / 90+)
 * onto a fresh "Aging" sheet.
 *
 * Edit the two column letters below to match your sheet, then run.
 */
const DUE_DATE_COLUMN = "D";   // <- column holding the due / scheduled date
const AMOUNT_COLUMN = "F";     // <- column holding the outstanding amount
const HEADER_ROWS = 1;         // <- how many header rows to skip

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const used = sheet.getUsedRange();
  if (!used) {
    console.log("Sheet is empty.");
    return;
  }

  const firstRow = used.getRowIndex() + HEADER_ROWS;
  const lastRow = used.getRowIndex() + used.getRowCount() - 1;
  if (lastRow < firstRow) {
    console.log("No data rows below the header.");
    return;
  }

  const dates = sheet
    .getRange(`${DUE_DATE_COLUMN}${firstRow + 1}:${DUE_DATE_COLUMN}${lastRow + 1}`)
    .getValues();
  const amounts = sheet
    .getRange(`${AMOUNT_COLUMN}${firstRow + 1}:${AMOUNT_COLUMN}${lastRow + 1}`)
    .getValues();

  const buckets = ["not due", "0-30", "31-60", "61-90", "90+"];
  const totals = new Map<string, { count: number; amount: number }>();
  buckets.forEach((b) => totals.set(b, { count: 0, amount: 0 }));

  // getValues() hands back dates as Excel serial numbers, so "today" has to be
  // one too. Day 0 of the 1900 system is 1899-12-30 -- two days before
  // 1900-01-01 rather than one, because Excel wrongly counts 1900 as a leap
  // year. Every date this business deals with is well past that, so the plain
  // offset is exact.
  const EXCEL_DAY_ZERO = Date.UTC(1899, 11, 30);
  const MS_PER_DAY = 86400000;
  const now = new Date();
  const todaySerial = Math.round(
    (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - EXCEL_DAY_ZERO) / MS_PER_DAY
  );

  let skipped = 0;

  for (let i = 0; i < dates.length; i++) {
    const serial = Number(dates[i][0]);
    const amount = Number(amounts[i][0]);
    if (!serial || isNaN(serial) || isNaN(amount) || amount === 0) {
      skipped++;
      continue;
    }

    const ageDays = todaySerial - serial;
    let bucket: string;
    if (ageDays <= 0) bucket = "not due";
    else if (ageDays <= 30) bucket = "0-30";
    else if (ageDays <= 60) bucket = "31-60";
    else if (ageDays <= 90) bucket = "61-90";
    else bucket = "90+";

    const t = totals.get(bucket)!;
    t.count++;
    t.amount += amount;
  }

  const existing = workbook.getWorksheet("Aging");
  if (existing) existing.delete();
  const out = workbook.addWorksheet("Aging");

  const rows: (string | number)[][] = [["Bucket", "Orders", "Outstanding"]];
  let grandCount = 0;
  let grandAmount = 0;
  for (const b of buckets) {
    const t = totals.get(b)!;
    rows.push([b, t.count, t.amount]);
    grandCount += t.count;
    grandAmount += t.amount;
  }
  rows.push(["Total", grandCount, grandAmount]);

  const target = out.getRangeByIndexes(0, 0, rows.length, 3);
  target.setValues(rows);
  out.getRange("A1:C1").getFormat().getFont().setBold(true);
  out.getRange(`A${rows.length}:C${rows.length}`).getFormat().getFont().setBold(true);
  out.getRange("C2:C" + rows.length).setNumberFormat("# ##0");
  out.getRange("A:C").getFormat().autofitColumns();
  out.activate();

  console.log(`Aging built from ${grandCount} row(s); ${skipped} row(s) skipped.`);
}
