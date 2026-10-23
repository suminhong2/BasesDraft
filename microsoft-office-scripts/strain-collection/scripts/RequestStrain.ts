// Request a strain from the collection
// Reads section 4 of the Forms tab, copies the strain to the Requests tab with
// Returned? = "no", adds CGC / WormBase links, and @mentions the people in
// Settings > "Email people about new requests" in a comment, which emails them.
function main(workbook: ExcelScript.Workbook) {
  const result = named(workbook, "RQ_Result");
  try {
    const message = requestStrain(workbook, result);
    if (message) result.setValue("✓ " + message);   // empty: the result already holds an order link
    clearForm(workbook, ["RQ_Strain", "RQ_Name"]);
  } catch (e) {
    const reason = String(e).replace(/^Error: /, "");
    result.setValue("✗ " + reason);
    // Also fail the run, so the Automate pane / button shows the reason instead of "complete"
    throw new Error(reason);
  }
}

function requestStrain(workbook: ExcelScript.Workbook, result: ExcelScript.Range): string {
  const strainName = read(workbook, "RQ_Strain");
  const who = read(workbook, "RQ_Name");
  if (!strainName) throw new Error("Pick the strain to request.");
  if (!who) throw new Error("Type your name.");

  const strains = table(workbook, "StrainTable");
  const headers = strains.getHeaderRowRange().getValues()[0].map(h => String(h));
  const rows = strains.getRangeBetweenHeaderAndTotal().getValues();
  const row = rows.find(r => String(r[0]).trim().toUpperCase() === strainName.toUpperCase());
  if (!row) throw new Error(strainName + " is not in the Strains tab.");
  const get = (h: string) => String(row[headers.indexOf(h)]).trim();
  const source = resolveSource(workbook, strainName, get("Publicly available?"));

  // Not in a box: point to where it can be ordered instead
  if (!/box/i.test(get("Box location"))) {
    const order = orderUrl(source, strainName);
    if (order) {
      result.setFormula('=HYPERLINK("' + order + '","✗ ' + strainName + ' has no box, so it is not in our collection. Click to ' +
        (source === "CGC" ? "order it from the CGC" : "request it from NBRP") + '.")');
      return "";
    }
    throw new Error(strainName + " has no box, so it is not in our collection. " +
      (source === "Own lab" ? "Ask the person who made it." : "Ask the lab it came from."));
  }

  // Still out with someone? Look for the newest request of this strain
  const requests = table(workbook, "RequestTable");
  const reqHeaders = requests.getHeaderRowRange().getValues()[0].map(h => String(h));
  const reqRows = requests.getRangeBetweenHeaderAndTotal().getValues();
  const col = (h: string) => reqHeaders.indexOf(h);
  let last: (string | number | boolean)[] = [];
  for (const r of reqRows) {
    if (String(r[col("Strain number")]).trim().toUpperCase() !== strainName.toUpperCase()) continue;
    if (!last.length || Number(r[col("Request date")]) >= Number(last[col("Request date")])) last = r;
  }
  if (last.length && /^\s*no/i.test(String(last[col("Returned?")]))) {
    throw new Error(strainName + " was requested by " + last[col("Requested by")] + " on " +
      excelDateText(Number(last[col("Request date")])) + " and is not returned yet. " +
      "Change Returned? to yes on the Requests tab when it is back.");
  }

  // The request row: request info, the strain's columns, then the links
  const values: { [header: string]: string | number } = {
    "Requested by": who, "Request date": excelNow(), "Returned?": "no", "Source": source,
  };
  headers.forEach((h, i) => { if (values[h] === undefined) values[h] = row[i] as string | number; });
  appendRows(requests, [rowByHeaders(requests, values)]);

  const newRow = requests.getRangeBetweenHeaderAndTotal().getLastRow();
  const order = orderUrl(source, strainName);
  if (order) {
    newRow.getCell(0, col("Order link")).setFormula('=HYPERLINK("' + order + '","' +
      (source === "CGC" ? "Order on CGC" : "Request from NBRP") + '")');
  }
  if (source !== "Own lab") {
    newRow.getCell(0, col("WormBase")).setFormula('=HYPERLINK("https://wormbase.org/search/strain/' +
      encodeURIComponent(strainName) + '","WormBase")');
  }

  const note = notify(workbook, newRow.getCell(0, 0), who + " requested " + strainName +
    " (" + get("Alleles / arrays") + "), " + get("Box location") + ".");
  return strainName + " requested by " + who + ". " + note +
    " Change Returned? to yes on the Requests tab when you return it.";
}

// Comments that @mention someone send them an email (work / school accounts, shared workbook)
function notify(workbook: ExcelScript.Workbook, cell: ExcelScript.Range, text: string): string {
  const emails = read(workbook, "NotifyEmails").split(/[,;\s]+/).filter(e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  if (!emails.length) return "Nobody is set to be emailed (Settings tab).";
  try {
    const mentions = emails.map((email, id) => ({ email: email, id: id, name: email.split("@")[0] }));
    const tags = mentions.map(m => '<at id="' + m.id + '">' + m.name + "</at>").join(" ");
    workbook.addComment(cell, { mentions: mentions, richContent: tags + " " + text }, ExcelScript.ContentType.mention);
    return "Emailed " + emails.join(", ") + " through a comment.";
  } catch (e) {
    return "The request was saved, but the email comment failed: " + String(e).replace(/^Error: /, "");
  }
}

// Where to order a strain: what "Publicly available?" says, "Own lab" for our lab code + a number,
// otherwise "Other lab". (Office Scripts can't read the CGC website, unlike the Google version.)
function resolveSource(workbook: ExcelScript.Workbook, name: string, hint: string): string {
  if (/cgc/i.test(hint)) return "CGC";
  if (/nbrp/i.test(hint)) return "NBRP";
  const labCode = read(workbook, "LabCode");
  if (labCode && new RegExp("^" + labCode + "\\d+$", "i").test(name)) return "Own lab";
  return "Other lab";
}

function orderUrl(source: string, name: string): string {
  if (source === "CGC") return "https://cgc.umn.edu/strain/" + encodeURIComponent(name);
  if (source === "NBRP") return "https://shigen.nig.ac.jp/c.elegans/";
  return "";
}

function excelDateText(serial: number): string {
  if (!(serial > 0)) return "an unknown date";
  return new Date(Math.round((serial - 25569) * 86400000)).toISOString().slice(0, 10);
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
