// Create a new prefix for strains from another lab (e.g. SH for Sumin Hong)
// Reads section 3 of the Forms tab. Checks the prefix list and the strain names already used.
// Is it an official CGC lab code? Open the CGC link on the form before adding it.
function main(workbook: ExcelScript.Workbook) {
  const result = named(workbook, "NP_Result");
  try {
    result.setValue("✓ " + addPrefix(workbook));
    clearForm(workbook, ["NP_Prefix", "NP_Owner"]);
    named(workbook, "NP_Start").setValue(1);
  } catch (e) {
    const reason = String(e).replace(/^Error: /, "");
    result.setValue("✗ " + reason);
    // Also fail the run, so the Automate pane / button shows the reason instead of "complete"
    throw new Error(reason);
  }
}

function addPrefix(workbook: ExcelScript.Workbook): string {
  const prefix = read(workbook, "NP_Prefix").toUpperCase();
  const owner = read(workbook, "NP_Owner");
  const start = Number(read(workbook, "NP_Start")) || 1;
  if (!/^[A-Z]{1,5}$/.test(prefix)) throw new Error("Use 1 to 5 letters only, e.g. SH.");
  if (!owner) throw new Error("Say whose prefix it is, full name.");
  if (start < 1 || Math.floor(start) !== start) throw new Error("Start numbering at must be a whole number, 1 or more.");

  const prefixTable = table(workbook, "PrefixTable");
  const prefixes = columnValues(prefixTable, "Prefix").map(p => p.toUpperCase());
  const at = prefixes.indexOf(prefix);
  if (at !== -1) {
    throw new Error(prefix + " is already taken by " + (columnValues(prefixTable, "Lab / person")[at] || "someone") +
      ". Pick it in section 2 if you belong to it, or choose another prefix (e.g. add a middle initial).");
  }
  const names = columnValues(table(workbook, "StrainTable"), "Strain number");
  const re = new RegExp("^" + prefix + "(\\d+)$", "i");
  let used = 0;
  names.forEach(n => {
    const m = n.match(re);
    if (m) used = Math.max(used, parseInt(m[1], 10));
  });
  if (used) {
    throw new Error(prefix + " is already used in the Strains tab (up to " + prefix + used +
      "). Ask who owns it, or choose another prefix.");
  }
  appendRows(prefixTable, [rowByHeaders(prefixTable, {
    "Prefix": prefix, "Lab / person": owner, "Added": excelNow(), "Note": "", "Start numbering at": start,
  })]);
  return prefix + " added. The first strain will be " + prefix + start + ".";
}

// ---------- helpers (the same in every script) ----------
function named(workbook: ExcelScript.Workbook, name: string): ExcelScript.Range {
  const item = workbook.getNamedItem(name);
  if (!item) throw new Error('The name "' + name + '" is missing. Use the template workbook.');
  return item.getRange();
}

function table(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const t = workbook.getTable(name);
  if (!t) throw new Error('The table "' + name + '" is missing. Use the template workbook.');
  return t;
}

function column(t: ExcelScript.Table, header: string): ExcelScript.Range {
  const c = t.getColumnByName(header);
  if (!c) throw new Error('The column "' + header + '" is missing. Keep the column names as they are.');
  return c.getRangeBetweenHeaderAndTotal();
}

function read(workbook: ExcelScript.Workbook, name: string): string {
  const value = named(workbook, name).getValue();
  return value === null || value === undefined ? "" : String(value).trim();
}

function clearForm(workbook: ExcelScript.Workbook, names: string[]) {
  names.forEach(n => named(workbook, n).setValue(""));
}

function columnValues(t: ExcelScript.Table, header: string): string[] {
  return column(t, header).getValues().map(r => String(r[0]).trim());
}

// A table row as a list, in the table's column order
function rowByHeaders(table: ExcelScript.Table, values: { [header: string]: string | number }): (string | number)[] {
  const headers = table.getHeaderRowRange().getValues()[0].map(h => String(h));
  return headers.map(h => (values[h] === undefined ? "" : values[h]));
}

// Adds rows at the bottom of a table. A new table starts with one empty row: fill that first.
function appendRows(table: ExcelScript.Table, rows: (string | number)[][]) {
  const body = table.getRangeBetweenHeaderAndTotal();
  const first = body.getValues();
  if (table.getRowCount() === 1 && first[0].every(v => String(v) === "")) {
    body.setValues([rows[0]]);
    rows = rows.slice(1);
  }
  if (rows.length) table.addRows(-1, rows);
}

// Now as an Excel date number, with the time (so same-day rows still sort in order)
function excelNow(): number {
  const d = new Date();
  return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569;
}
