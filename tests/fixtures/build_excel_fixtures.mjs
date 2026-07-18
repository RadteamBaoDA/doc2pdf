import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = path.resolve("tests/fixtures/generated");
await fs.mkdir(outDir, { recursive: true });

function matrix(rows, cols, fn) {
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => fn(r, c)),
  );
}

function styleHeader(range) {
  range.format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function seedDataSheet(sheet, rows, cols, prefix = "Data") {
  const headers = Array.from({ length: cols }, (_, c) =>
    c === 0 ? "LEFT_SENTINEL" : c === cols - 1 ? "RIGHT_SENTINEL" : `${prefix}_${c + 1}`,
  );
  sheet.getRangeByIndexes(0, 0, 1, cols).values = [headers];
  sheet.getRangeByIndexes(1, 0, rows, cols).values = matrix(rows, cols, (r, c) => {
    if (c === 0) return `LEFT_${r + 1}`;
    if (c === cols - 1) return `RIGHT_${r + 1}`;
    return (r + 1) * (c + 1);
  });
  styleHeader(sheet.getRangeByIndexes(0, 0, 1, cols));
  sheet.getRangeByIndexes(0, 0, rows + 1, cols).format.borders = {
    preset: "inside",
    style: "thin",
    color: "#D9E2F3",
  };
  sheet.getRangeByIndexes(0, 0, rows + 1, cols).format.columnWidth = 12;
  sheet.getRangeByIndexes(0, 0, 1, cols).format.rowHeight = 28;
  sheet.freezePanes.freezeRows(1);
}

async function writeWorkbook(name, build) {
  const workbook = Workbook.create();
  await build(workbook);
  const result = await SpreadsheetFile.exportXlsx(workbook);
  await result.save(path.join(outDir, name));
}

await writeWorkbook("excel_printareas_shapes.xlsx", async (wb) => {
  const sheet = wb.worksheets.add("PrintAreas");
  sheet.getRange("B3:F6").values = [
    ["LEFT_SENTINEL", "Merged", "Unicode", "Wrap", "Area 1"],
    ["区域一", "内容", "Đường biên trái", "This deliberately wraps across a narrow region", 101],
    ["A3", "B3", "C3", "D3", 103],
    ["A4", "B4", "C4", "D4", 104],
  ];
  sheet.getRange("H20:J21").values = [
    ["RIGHT_SENTINEL", "Area 2", "Ω"],
    ["visible-right", "second-area", 202],
  ];
  sheet.mergeCells("C3:D3");
  sheet.getRange("B3:F3").format = { fill: "#5B9BD5", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("B4:F6").format.wrapText = true;
  const chart = sheet.charts.add("bar", sheet.getRange("I20:J21"));
  chart.title = "Visible chart shape";
  chart.setPosition("L3", "T18");
  sheet.freezePanes.freezeRows(3);
});

await writeWorkbook("excel_chunks.xlsx", async (wb) => {
  const sheet = wb.worksheets.add("Chunks");
  seedDataSheet(sheet, 120, 40, "CHUNK");
  sheet.getRange("A1:AN1").format.fill = "#1F4E78";
  sheet.getRange("A2").values = [["Chunk 1 starts here"]];
  sheet.getRange("A121").values = [["Chunk 12 ends here"]];
});

await writeWorkbook("excel_max_width.xlsx", async (wb) => {
  const sheet = wb.worksheets.add("MaxWidth");
  seedDataSheet(sheet, 8, 2048, "WIDE");
  sheet.getRange("A1:B2").format.columnWidth = 24;
  sheet.getRange("A1").values = [["LEFT_SENTINEL - do not lose"]];
  sheet.getRangeByIndexes(0, 2047, 2, 1).format.columnWidth = 24;
  sheet.getRangeByIndexes(0, 2047, 1, 1).values = [["RIGHT_SENTINEL - do not lose"]];
});

await writeWorkbook("excel_oversized_error.xlsx", async (wb) => {
  const normal = wb.worksheets.add("Normal");
  seedDataSheet(normal, 5, 12, "NORMAL");
  const wide = wb.worksheets.add("Oversized");
  seedDataSheet(wide, 3, 512, "OVERSIZED");
  wide.getRange("A1:B4").format.columnWidth = 90;
  wide.getRangeByIndexes(0, 511, 1, 1).values = [["RIGHT_SENTINEL"]];
  const max = wb.worksheets.add("MaxColumns");
  const maxCols = 16384;
  max.getRangeByIndexes(0, 0, 2, maxCols).values = [
    Array.from({ length: maxCols }, (_, c) => c === 0 ? "LEFT_SENTINEL" : c === maxCols - 1 ? "RIGHT_SENTINEL" : `MAX_${c + 1}`),
    Array.from({ length: maxCols }, (_, c) => c + 1),
  ];
  max.getRangeByIndexes(0, 0, 1, maxCols).format = { fill: "#7F6000", font: { bold: true, color: "#FFFFFF" } };
});

await writeWorkbook("excel_oversized_skip.xlsx", async (wb) => {
  const normal = wb.worksheets.add("KeepMe");
  seedDataSheet(normal, 10, 12, "KEEP");
  const wide = wb.worksheets.add("SkipMe");
  seedDataSheet(wide, 3, 512, "SKIP");
  wide.getRange("A1:Z4").format.columnWidth = 90;
});

await writeWorkbook("excel_empty_shapes.xlsx", async (wb) => {
  wb.worksheets.add("EmptyVisible");
  const shapes = wb.worksheets.add("ShapesOnly");
  shapes.getRange("A1:B3").values = [["SHAPE_LEFT", ""], ["", ""], ["", "SHAPE_RIGHT"]];
  const chart = shapes.charts.add("bar", shapes.getRange("A1:B3"));
  chart.title = "Shapes-only content anchor";
  chart.setPosition("D4", "L18");
  wb.worksheets.add("HiddenCandidate");
});

await writeWorkbook("excel_empty_workbook.xlsx", async (wb) => {
  wb.worksheets.add("EmptyVisible");
  wb.worksheets.add("HiddenCandidate");
});

await writeWorkbook("excel_orientation_margin.xlsx", async (wb) => {
  for (const name of ["Portrait", "Landscape"]) {
    const sheet = wb.worksheets.add(name);
    seedDataSheet(sheet, 25, name === "Portrait" ? 8 : 18, name.toUpperCase());
    sheet.getRange("A1").values = [[`${name} LEFT_SENTINEL`]];
    sheet.getRangeByIndexes(0, name === "Portrait" ? 7 : 17, 1, 1).values = [[`${name} RIGHT_SENTINEL`]];
    sheet.getRange("A2:A5").format.wrapText = true;
    sheet.mergeCells("C2:E2");
  }
});

await writeWorkbook("excel_quality_pagination.xlsx", async (wb) => {
  const wide = wb.worksheets.add("WideOnly");
  seedDataSheet(wide, 12, 64, "QUALITY_WIDE");
  wide.getRange("A1").values = [["WIDE_LEFT_SENTINEL"]];
  wide.getRangeByIndexes(0, 63, 1, 1).values = [["WIDE_RIGHT_SENTINEL"]];

  const tall = wb.worksheets.add("TallOnly");
  seedDataSheet(tall, 180, 8, "QUALITY_TALL");
  tall.getRange("A1").values = [["TALL_TOP_SENTINEL"]];
  tall.getRange("A181").values = [["TALL_BOTTOM_SENTINEL"]];
  tall.getRange("A2:H181").format.rowHeight = 30;

  const both = wb.worksheets.add("WideTall");
  seedDataSheet(both, 80, 48, "QUALITY_2D");
  both.getRange("A1").values = [["2D_LEFT_TOP_SENTINEL"]];
  both.getRangeByIndexes(0, 47, 1, 1).values = [["2D_RIGHT_SENTINEL"]];
  both.getRange("A81").values = [["2D_BOTTOM_SENTINEL"]];
  both.getRange("A2:AV81").format.rowHeight = 36;
  both.mergeCells("C2:E2");
  both.getRange("C2").values = [["MERGED_SENTINEL"]];
  const chart = both.charts.add("bar", both.getRange("A3:B12"));
  chart.title = "QUALITY_CHART_SENTINEL";
  chart.setPosition("B60", "H78");
});

await writeWorkbook("excel_paper_forms.xlsx", async (wb) => {
  const bands = [
    ["A4Portrait", 7.0], ["LetterPortrait", 7.4], ["B4Portrait", 8.2],
    ["TabloidPortrait", 9.2], ["A3Portrait", 10.4], ["LetterLandscape", 10.8],
    ["LegalLandscape", 13.2], ["LedgerLandscape", 16.2],
    ["ArchCLandscape", 21.2], ["ArchDLandscape", 33.0], ["ArchELandscape", 43.0],
  ];
  for (const [name, targetWidth] of bands) {
    const sheet = wb.worksheets.add(name);
    sheet.getRange("A1:B5").values = [
      [`${name} LEFT_SENTINEL`, "band target width"],
      ["visible content", targetWidth],
      ["Unicode Ω Đ 一", "right edge"],
      ["formula source", 100],
      ["RIGHT_SENTINEL", targetWidth],
    ];
    sheet.getRange("A1:B1").format = { fill: "#203864", font: { bold: true, color: "#FFFFFF" } };
    // Width in Excel character units; the COM seeder records the measured width
    // and applies orientation/form enums for deterministic acceptance runs.
    sheet.getRange("A1:A5").format.columnWidth = Math.max(12, targetWidth * 5);
    sheet.getRange("B1:B5").format.columnWidth = 16;
  }
});

console.log(`Generated Excel fixture sources in ${outDir}`);
