import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const samplesDir = path.join(repoRoot, "samples");
const sourceDir = path.join(samplesDir, "source");
const expectedDir = path.join(samplesDir, "expected");
const checkedDir = path.join(samplesDir, "checked");
const force = process.argv.includes("--force");
const render = process.argv.includes("--render");

const months = [
  { month: "2026-02", yymm: "2602", daysInMonth: 28, missingDays: [3] },
  { month: "2026-04", yymm: "2604", daysInMonth: 30, missingDays: [3] },
  { month: "2026-05", yymm: "2605", daysInMonth: 31, missingDays: [2] },
];

const branchBusinessMap = {
  BR01: ["INQ", "RSP", "REP", "BIL"],
  BR02: ["INQ", "DEL", "APP"],
  BR03: ["INQ", "RSP", "CON"],
  BR04: ["INQ", "RSP", "REP", "DEL", "BIL"],
  BR05: ["INQ", "APP", "CON", "OTH"],
  BR06: ["RSP", "DEL", "OTH"],
};

function csvEscape(value) {
  const text = value === null || value === undefined ? "" : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function csvLine(values) {
  return values.map(csvEscape).join(",");
}

function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines.shift().split(",");
  return lines.map((line) => {
    const values = [];
    let current = "";
    let quoted = false;
    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];
      if (ch === '"' && quoted && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else if (ch === '"') {
        quoted = !quoted;
      } else if (ch === "," && !quoted) {
        values.push(current);
        current = "";
      } else {
        current += ch;
      }
    }
    values.push(current);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

async function readCsv(fileName) {
  return parseCsv(await fs.readFile(path.join(sourceDir, fileName), "utf8"));
}

function colName(index) {
  let n = index;
  let name = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function rangeAddress(row, col, rowCount, colCount) {
  const start = `${colName(col)}${row}`;
  const end = `${colName(col + colCount - 1)}${row + rowCount - 1}`;
  return `${start}:${end}`;
}

function setValues(sheet, row, col, values) {
  sheet.getRange(rangeAddress(row, col, values.length, values[0].length)).values = values;
}

function formatTable(sheet, row, col, rowCount, colCount) {
  const table = sheet.getRange(rangeAddress(row, col, rowCount, colCount));
  table.format = {
    font: { name: "Yu Gothic", size: 10 },
    borders: { preset: "all", style: "thin", color: "#D1D5DB" },
    verticalAlignment: "center",
  };
  const header = sheet.getRange(rangeAddress(row, col, 1, colCount));
  header.format = {
    fill: "#DDEBF7",
    font: { name: "Yu Gothic", size: 10, bold: true, color: "#1F2937" },
    borders: { preset: "all", style: "thin", color: "#9CA3AF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function dayFromDate(dateText) {
  return Number(dateText.slice(8, 10));
}

function monthMeta(month) {
  return months.find((item) => item.month === month);
}

function buildLayoutRows(branches, businesses) {
  const branchByCode = Object.fromEntries(branches.map((branch) => [branch.branch_code, branch]));
  const businessByCode = Object.fromEntries(businesses.map((business) => [business.business_code, business]));
  const rows = [];
  const subtotalSeen = new Set();

  for (const branch of branches) {
    for (const businessCode of branchBusinessMap[branch.branch_code]) {
      const business = businessByCode[businessCode];
      rows.push({
        row_type: "DETAIL",
        subtotal_group: branch.subtotal_group,
        region: branch.region,
        branch_code: branch.branch_code,
        branch_name: branch.branch_name,
        business_code: business.business_code,
        business_name: business.business_name,
        note: "",
      });
    }
    if (!subtotalSeen.has(branch.subtotal_group)) {
      subtotalSeen.add(branch.subtotal_group);
      const label = branch.branch_size === "large"
        ? `${branch.branch_name} 小計`
        : `${branch.subtotal_group} 小計`;
      rows.push({
        row_type: "SUBTOTAL",
        subtotal_group: branch.subtotal_group,
        region: branch.region,
        branch_code: "",
        branch_name: label,
        business_code: "",
        business_name: "",
        note: "小計行。入力キーではありません。",
      });
    }
  }
  return rows;
}

function computeExpectedRows(layoutRows, records, month) {
  const meta = monthMeta(month);
  const rows = layoutRows.map((row) => ({
    ...row,
    month,
    days: Array.from({ length: 31 }, () => ""),
  }));

  const detailIndex = new Map(rows
    .filter((row) => row.row_type === "DETAIL")
    .map((row) => [`${row.branch_code}|${row.business_code}`, row]));

  for (const record of records.filter((item) => item.month === month && item.status === "ok")) {
    const key = `${record.branch_code}|${record.business_code}`;
    const target = detailIndex.get(key);
    if (!target) continue;
    const day = dayFromDate(record.date);
    const value = Number(record.complaint_count);
    target.days[day - 1] = (Number(target.days[day - 1]) || 0) + value;
  }

  for (const row of rows.filter((item) => item.row_type === "SUBTOTAL")) {
    const detailRows = rows.filter((item) => item.row_type === "DETAIL" && item.subtotal_group === row.subtotal_group);
    for (let i = 0; i < meta.daysInMonth; i += 1) {
      const total = detailRows.reduce((sum, detail) => sum + (Number(detail.days[i]) || 0), 0);
      row.days[i] = total === 0 ? "" : total;
    }
  }

  return rows;
}

function expectedMonthlyCsvRows(layoutRows, records) {
  const header = [
    "month", "row_type", "subtotal_group", "region", "branch_code", "branch_name",
    "business_code", "business_name",
    ...Array.from({ length: 31 }, (_, index) => `day_${String(index + 1).padStart(2, "0")}`),
    "month_total", "note",
  ];
  const lines = [csvLine(header)];
  for (const { month, daysInMonth } of months) {
    for (const row of computeExpectedRows(layoutRows, records, month)) {
      const dayValues = row.days.map((value, index) => (index < daysInMonth ? value : ""));
      const monthTotal = dayValues.reduce((sum, value) => sum + (Number(value) || 0), 0);
      lines.push(csvLine([
        month, row.row_type, row.subtotal_group, row.region, row.branch_code, row.branch_name,
        row.business_code, row.business_name, ...dayValues, monthTotal || "", row.note,
      ]));
    }
  }
  return `${lines.join("\n")}\n`;
}

function anomalyCsvRows(records) {
  const header = ["month", "date", "file_name", "subfolder", "row_no", "issue_type", "expected_behavior"];
  const lines = [csvLine(header)];
  for (const record of records.filter((item) => item.status !== "ok")) {
    lines.push(csvLine([
      record.month, record.date, record.file_name, record.subfolder, record.row_no,
      record.status, record.note,
    ]));
  }
  for (const meta of months) {
    for (const day of meta.missingDays) {
      const yymmdd = `${meta.yymm}${String(day).padStart(2, "0")}`;
      lines.push(csvLine([
        meta.month, `${meta.month}-${String(day).padStart(2, "0")}`,
        `クレーム集計${yymmdd}.xlsx`, "", "", "missing_day",
        "転記なし。日付列は空欄のまま",
      ]));
    }
    for (let day = meta.daysInMonth + 1; day <= 31; day += 1) {
      lines.push(csvLine([
        meta.month, "", "", "", "", "out_of_month_day_column",
        `${meta.month} の ${day} 日列は空欄のまま`,
      ]));
    }
  }
  return `${lines.join("\n")}\n`;
}

async function verifyAndExport(workbook, outputPath, renderRangesBySheet) {
  if (!force) {
    try {
      await fs.access(outputPath);
      return;
    } catch {
      // Missing files are generated below.
    }
  }
  workbook.recalculate();
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "formula error scan",
  });
  const errorMatches = errors.ndjson
    .trim()
    .split(/\r?\n/)
    .filter((line) => line.includes('"kind":"match"'));
  if (errorMatches.length > 0) {
    throw new Error(`Formula errors found before export for ${outputPath}:\n${errors.ndjson}`);
  }
  if (render) {
    for (const [sheetName, range] of Object.entries(renderRangesBySheet)) {
      await workbook.render({ sheetName, range, scale: 1 });
    }
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await output.save(outputPath);
}

async function buildDailyWorkbook(records, outputPath) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("日次集計");
  const headers = ["処理区分", "支店コード", "支店名", "業務コード", "業務名", "クレーム件数", "備考"];
  const values = [
    headers,
    ...records.map((record) => [
      record.status,
      record.branch_code,
      record.branch_name,
      record.business_code,
      record.business_name,
      record.complaint_count,
      record.note,
    ]),
  ];
  setValues(sheet, 1, 1, values);
  formatTable(sheet, 1, 1, values.length, headers.length);
  sheet.getRange("A1:G1").format.fill = "#E2F0D9";
  sheet.getRange("A:G").format.autofitColumns();
  await verifyAndExport(workbook, outputPath, { "日次集計": `A1:G${Math.max(values.length, 8)}` });
}

function monthlyHeaders() {
  return [
    "行種別", "小計グループ", "地域", "支店コード", "支店名", "業務コード", "業務名",
    ...Array.from({ length: 31 }, (_, index) => String(index + 1)),
    "月合計", "備考",
  ];
}

function monthlyRowValues(row, dayValues, includeExpected) {
  const total = dayValues.reduce((sum, value) => sum + (Number(value) || 0), 0);
  return [
    row.row_type, row.subtotal_group, row.region, row.branch_code, row.branch_name,
    row.business_code, row.business_name,
    ...dayValues.map((value) => (includeExpected ? value : "")),
    includeExpected ? (total || "") : "",
    row.note,
  ];
}

function addMonthlySheet(workbook, sheetName, month, rows, includeExpected) {
  const sheet = workbook.worksheets.add(sheetName);
  const headers = monthlyHeaders();
  sheet.getRange("A1").values = [[`${sheetName} ${month}`]];
  sheet.getRange("A1:AN1").format = {
    fill: "#1F4E78",
    font: { name: "Yu Gothic", size: 12, bold: true, color: "#FFFFFF" },
  };
  setValues(sheet, 3, 1, [headers]);
  const values = rows.map((row) => monthlyRowValues(row, row.days, includeExpected));
  setValues(sheet, 4, 1, values);
  formatTable(sheet, 3, 1, values.length + 1, headers.length);
  const firstDataRow = 4;
  const lastDataRow = firstDataRow + values.length - 1;
  const firstDayCol = 8;
  const totalCol = 39;
  const detailOrSubtotalRows = rows.map((row, index) => ({ ...row, excelRow: firstDataRow + index }));

  if (!includeExpected) {
    for (const row of detailOrSubtotalRows) {
      const totalCell = sheet.getRange(`${colName(totalCol)}${row.excelRow}`);
      totalCell.formulas = [[`=SUM(H${row.excelRow}:AL${row.excelRow})`]];
      if (row.row_type === "SUBTOTAL") {
        for (let dayIndex = 0; dayIndex < 31; dayIndex += 1) {
          const column = colName(firstDayCol + dayIndex);
          sheet.getRange(`${column}${row.excelRow}`).formulas = [[`=SUMIFS(${column}$${firstDataRow}:${column}$${lastDataRow},$B$${firstDataRow}:$B$${lastDataRow},$B${row.excelRow},$A$${firstDataRow}:$A$${lastDataRow},"DETAIL")`]];
        }
      }
    }
  }

  for (const row of detailOrSubtotalRows.filter((item) => item.row_type === "SUBTOTAL")) {
    sheet.getRange(`A${row.excelRow}:AN${row.excelRow}`).format.fill = "#FFF2CC";
    sheet.getRange(`A${row.excelRow}:AN${row.excelRow}`).format.font = { name: "Yu Gothic", size: 10, bold: true };
  }

  const meta = monthMeta(month);
  if (meta.daysInMonth < 31) {
    const startCol = colName(firstDayCol + meta.daysInMonth);
    sheet.getRange(`${startCol}3:AL${lastDataRow}`).format.fill = "#E7E6E6";
  }
  sheet.getRange("A:AN").format.autofitColumns();
  return sheet;
}

function addMasterSheet(workbook, branches, businesses, layoutRows) {
  const sheet = workbook.worksheets.add("マスタ");
  setValues(sheet, 1, 1, [["支店マスタ"]]);
  setValues(sheet, 2, 1, [
    ["支店コード", "支店名", "地域", "支店区分", "小計グループ"],
    ...branches.map((branch) => [
      branch.branch_code,
      branch.branch_name,
      branch.region,
      branch.branch_size === "large" ? "大支店" : "小支店",
      branch.subtotal_group,
    ]),
  ]);
  formatTable(sheet, 2, 1, branches.length + 1, 5);

  setValues(sheet, 11, 1, [["業務マスタ"]]);
  setValues(sheet, 12, 1, [
    ["業務コード", "業務名"],
    ...businesses.map((business) => [business.business_code, business.business_name]),
  ]);
  formatTable(sheet, 12, 1, businesses.length + 1, 2);

  const detailRows = layoutRows.filter((row) => row.row_type === "DETAIL");
  setValues(sheet, 23, 1, [["支店・業務対応表"]]);
  setValues(sheet, 24, 1, [
    ["支店コード", "支店名", "業務コード", "業務名", "小計グループ"],
    ...detailRows.map((row) => [row.branch_code, row.branch_name, row.business_code, row.business_name, row.subtotal_group]),
  ]);
  formatTable(sheet, 24, 1, detailRows.length + 1, 5);
  sheet.getRange("A:E").format.autofitColumns();
}

function addAnomalySheet(workbook, anomalies, month) {
  const sheet = workbook.worksheets.add("異常系");
  const headers = ["month", "date", "file_name", "subfolder", "row_no", "issue_type", "expected_behavior"];
  const displayHeaders = ["対象月", "日付", "ファイル名", "サブフォルダ", "行番号", "異常種別", "期待動作"];
  const rows = anomalies.filter((row) => row.month === month);
  setValues(sheet, 1, 1, [displayHeaders, ...rows.map((row) => headers.map((header) => row[header]))]);
  formatTable(sheet, 1, 1, rows.length + 1, headers.length);
  sheet.getRange("A:G").format.autofitColumns();
}

async function buildMonthlyWorkbook(month, rows, branches, businesses, layoutRows, anomalies, outputPath) {
  const workbook = Workbook.create();
  addMonthlySheet(workbook, "月次集計", month, rows, false);
  addMonthlySheet(workbook, "期待結果", month, rows, true);
  addMasterSheet(workbook, branches, businesses, layoutRows);
  addAnomalySheet(workbook, anomalies, month);
  await verifyAndExport(workbook, outputPath, {
    "月次集計": "A1:AN35",
    "期待結果": "A1:AN35",
    "マスタ": "A1:E48",
    "異常系": "A1:G20",
  });
}

async function main() {
  const branches = await readCsv("branches.csv");
  const businesses = await readCsv("business-lines.csv");
  const records = await readCsv("daily-records.csv");
  const layoutRows = buildLayoutRows(branches, businesses);

  await fs.mkdir(expectedDir, { recursive: true });
  await fs.writeFile(path.join(expectedDir, "monthly-expected.csv"), expectedMonthlyCsvRows(layoutRows, records), "utf8");
  await fs.writeFile(path.join(expectedDir, "anomaly-expected.csv"), anomalyCsvRows(records), "utf8");

  const anomalies = parseCsv(await fs.readFile(path.join(expectedDir, "anomaly-expected.csv"), "utf8"));

  if (force) {
    await fs.rm(path.join(checkedDir, "daily"), { recursive: true, force: true });
    await fs.rm(path.join(checkedDir, "monthly"), { recursive: true, force: true });
  }

  const groups = new Map();
  for (const record of records) {
    const key = `${record.month}|${record.subfolder}|${record.file_name}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(record);
  }

  for (const [key, groupRecords] of groups.entries()) {
    const [month, subfolder, fileName] = key.split("|");
    const outputPath = path.join(checkedDir, "daily", monthMeta(month).yymm, subfolder, fileName);
    await buildDailyWorkbook(groupRecords, outputPath);
  }

  for (const meta of months) {
    const rows = computeExpectedRows(layoutRows, records, meta.month);
    const outputPath = path.join(checkedDir, "monthly", `月次クレーム集計${meta.yymm}.xlsx`);
    await buildMonthlyWorkbook(meta.month, rows, branches, businesses, layoutRows, anomalies, outputPath);
  }

  console.log(`日次workbook ${groups.size}件、月次workbook ${months.length}件を生成しました。`);
}

try {
  await main();
  process.exit(0);
} catch (error) {
  console.error(error);
  process.exit(1);
}
