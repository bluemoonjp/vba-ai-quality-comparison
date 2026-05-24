import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const workbookDir = path.join(root, "workbooks");
const dailyDir = path.join(workbookDir, "daily", "2605");

function parseCsv(text) {
  const [headerLine, ...lines] = text.trim().split(/\r?\n/);
  const headers = headerLine.split(",");
  return lines.map((line) => {
    const values = line.split(",");
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

async function readCsv(name) {
  const text = await fs.readFile(path.join(here, name), "utf8");
  return parseCsv(text);
}

function setHeader(range) {
  range.format = {
    fill: { color: "#D9EAF7" },
    font: { bold: true, color: "#1F2933" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

async function saveWorkbook(workbook, filePath) {
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(filePath);
}

async function buildMonthly(branches, businesses) {
  const workbook = Workbook.create();
  const monthly = workbook.worksheets.add("月次集計");
  const master = workbook.worksheets.add("マスタ");
  const expected = workbook.worksheets.add("期待結果");

  const dayHeaders = Array.from({ length: 31 }, (_, index) => `${index + 1}日`);
  const headers = ["支店コード", "支店名", "業務コード", "業務名", ...dayHeaders, "月合計"];
  monthly.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  setHeader(monthly.getRangeByIndexes(0, 0, 1, headers.length));

  const detailRows = [];
  for (const branch of branches) {
    for (const business of businesses) {
      detailRows.push([
        branch.branch_code,
        branch.branch_name,
        business.business_code,
        business.business_name,
        ...Array(31).fill(null),
        null,
      ]);
    }
  }
  monthly.getRangeByIndexes(1, 0, detailRows.length, headers.length).values = detailRows;
  monthly.getRangeByIndexes(0, 0, detailRows.length + 1, headers.length).format.autofitColumns();
  monthly.freezePanes.freezeRows(1);

  master.getRange("A1:C1").values = [["支店コード", "支店名", "地域"]];
  setHeader(master.getRange("A1:C1"));
  master.getRangeByIndexes(1, 0, branches.length, 3).values = branches.map((branch) => [
    branch.branch_code,
    branch.branch_name,
    branch.area,
  ]);
  master.getRange("E1:F1").values = [["業務コード", "業務名"]];
  setHeader(master.getRange("E1:F1"));
  master.getRangeByIndexes(1, 4, businesses.length, 2).values = businesses.map((business) => [
    business.business_code,
    business.business_name,
  ]);
  master.getUsedRange().format.autofitColumns();

  const expectedCsv = await fs.readFile(path.join(root, "expected", "monthly-expected.csv"), "utf8");
  const expectedRows = expectedCsv.trim().split(/\r?\n/).map((line) => line.split(","));
  expected.getRangeByIndexes(0, 0, expectedRows.length, expectedRows[0].length).values = expectedRows;
  setHeader(expected.getRangeByIndexes(0, 0, 1, expectedRows[0].length));
  expected.getUsedRange().format.autofitColumns();

  await saveWorkbook(workbook, path.join(workbookDir, "月次クレーム集計2605.xlsx"));
}

async function buildDaily(records) {
  const grouped = new Map();
  for (const record of records) {
    if (!grouped.has(record.date)) grouped.set(record.date, []);
    grouped.get(record.date).push(record);
  }

  for (const [date, rows] of grouped) {
    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add("日次集計");
    const headers = ["支店コード", "支店名", "業務コード", "業務名", "クレーム件数"];
    sheet.getRange("A1:E1").values = [headers];
    setHeader(sheet.getRange("A1:E1"));
    sheet.getRangeByIndexes(1, 0, rows.length, 5).values = rows.map((row) => [
      row.branch_code,
      row.branch_name,
      row.business_code,
      row.business_name,
      Number(row.claim_count),
    ]);
    sheet.getUsedRange().format.autofitColumns();
    sheet.freezePanes.freezeRows(1);
    await saveWorkbook(workbook, path.join(dailyDir, `クレーム集計${date}.xlsx`));
  }
}

await fs.mkdir(dailyDir, { recursive: true });
const branches = await readCsv("branches.csv");
const businesses = await readCsv("business-lines.csv");
const records = await readCsv("daily-records.csv");
await buildMonthly(branches, businesses);
await buildDaily(records);
