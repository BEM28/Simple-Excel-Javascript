const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet1");

const mergeCellOptionsCenter = {
  horizontal: "center",
  vertical: "middle",
};
const mergeCellOptionsLeft = {
  horizontal: "left",
  vertical: "middle",
};

worksheet.getCell("A1").value = "Item In Recap Report";
worksheet.getCell("A1").alignment = mergeCellOptionsCenter;
worksheet.getCell("A2").value = "Term : Desember-2023";
worksheet.getCell("A2").alignment = mergeCellOptionsCenter;
worksheet.getCell("A3").value = "Warehouse : All";
worksheet.getCell("A3").alignment = mergeCellOptionsCenter;
worksheet.getCell("A4").value = "Print Date : 02-January-2024";
worksheet.getCell("A4").alignment = mergeCellOptionsCenter;
worksheet.addRow([]);
worksheet.addRow([]);
// Menambahkan header pertama
const headers = ["", "", "", "", "Desember", ""];
const headers2 = [
  "No.",
  "Item Code",
  "Item Name",
  "Unit",
  1,
  2,
  3,
  4,
  5,
  6,
  7,
  8,
  9,
  10,
  11,
  12,
  13,
  14,
  15,
  16,
  17,
  18,
  19,
  20,
  21,
  22,
  23,
  24,
  25,
  26,
  27,
  28,
  29,
  30,
  31,
  "Total",
  "Harga Total",
  "Harga Satuan",
];
// Menambahkan data
const data = [
  [
    1,
    "ABC123",
    "Product A",
    "Pcs",
    10,
    20,
    30,
    40,
    50,
    60,
    70,
    80,
    90,
    100,
    110,
    120,
    130,
    140,
    150,
    160,
    170,
    180,
    190,
    200,
    210,
    220,
    230,
    240,
    250,
    260,
    270,
    280,
    290,
    300,
    315,
    7500,
    225000,
    3000,
  ],
  [
    2,
    "XYZ789",
    "Product B",
    "Kg",
    5,
    10,
    15,
    20,
    25,
    30,
    35,
    40,
    45,
    50,
    55,
    60,
    65,
    70,
    75,
    80,
    85,
    90,
    95,
    100,
    105,
    110,
    115,
    120,
    125,
    130,
    135,
    140,
    145,
    150,
    160,
    3750,
    150000,
    2500,
  ],
];
// menambahkan data ke worksheet
const header1Row = worksheet.addRow(headers);
const header2Row = worksheet.addRow(headers2);
data.forEach((rowData, rowIndex) => {
  const dataRow = worksheet.addRow(rowData);
  dataRow.eachCell((cell, index) => {
    cell.border = {
      top: { style: index === 0 ? "thin" : "none" },
      right: { style: index === 36 ? "thin" : "none" },
      bottom: {
        style:
          rowIndex === data.length - 1 && index !== 37 && index !== 38
            ? "thin"
            : "none",
      },
    };

    if (index == 1) {
      cell.alignment = mergeCellOptionsLeft;
    }

    if (index >= 5 && index <= 36 && index % 2 === 1) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "EFEFEF" },
      };
    }
  });
});

let totalLength = data.length + 9;
let hargaTotalLength = data.length + 10;

const totalText = "Total Data : " + data.length;
const hargaTotalFormula = `SUM(AK9:AK${hargaTotalLength - 1})`;
const hargaTotal = { formula: hargaTotalFormula };

worksheet.getCell("A" + totalLength).value = totalText;
worksheet.getCell("AK" + hargaTotalLength).value = hargaTotal;
// style
worksheet.mergeCells("E7:AJ7");
worksheet.mergeCells("A1:AI1");
worksheet.mergeCells("A2:AI2");
worksheet.mergeCells("A3:AI3");
worksheet.mergeCells("A4:AI4");

worksheet.getColumn(1).width = 4;
worksheet.getColumn(2).width = 9;
worksheet.getColumn(3).width = 31;
worksheet.getColumn(4).width = 6;
for (let col = 5; col <= 35; col++) {
  worksheet.getColumn(col).width = 4;
}
worksheet.getColumn(36).width = 6;
worksheet.getColumn(37).width = 14;
worksheet.getColumn(38).width = 14;

header1Row.eachCell((cell, index) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "C0C0C0" },
  };
  if (index == 36) {
    cell.border = {
      right: { style: "thin" },
      top: { style: "thin" },
      bottom: { style: "thin" },
    };
  } else {
    cell.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
    };
  }
});

header2Row.eachCell((cell, index) => {
  if (index == 37 || index == 38) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF" },
    };
  } else {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "C0C0C0" },
    };
  }
  if (index >= 5 && index <= 38) {
    cell.alignment = mergeCellOptionsCenter;
  }

  if (index == 36) {
    cell.border = {
      right: { style: "thin" },
      top: { style: "thin" },
      bottom: { style: "thin" },
    };
  } else if (index >= 1 && index <= 35) {
    cell.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
    };
  }
});
worksheet.getCell("E7").alignment = mergeCellOptionsCenter;
worksheet.getCell("A1").alignment = mergeCellOptionsCenter;
worksheet.getCell("A2").alignment = mergeCellOptionsCenter;
worksheet.getCell("A3").alignment = mergeCellOptionsCenter;
worksheet.getCell("A4").alignment = mergeCellOptionsCenter;

worksheet.views = [
  { state: "frozen", xSplit: 4, ySplit: 8, activeCell: "D1" },
  { showGridLines: false },
];

// Menyimpan Workbook ke File Excel
workbook.xlsx
  .writeFile("in_recap.xlsx")
  .then(() => {
    console.log("File Excel berhasil disimpan.");
  })
  .catch((error) => {
    console.error("Gagal menyimpan file Excel:", error);
  });
