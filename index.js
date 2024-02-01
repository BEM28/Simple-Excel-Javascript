import { Workbook } from "exceljs";

const workbook = new Workbook();
const worksheet = workbook.addWorksheet("Sheet1");

worksheet.addRow(["Stock Opname Report"]);
worksheet.addRow(["Term : 01-December-2023 - 31-December-2023"]);
worksheet.addRow(["Print Date : 02-January-2024"]);
worksheet.addRow([]);

// Menambahkan header pertama
const headers = [
  "no",
  "",
  "",
  "",
  "",
  "",
  "Code",
  "Date",
  "Warehouse",
  "status",
  "",
  "",
];
const headers2 = [
  "",
  "",
  "",
  "",
  "",
  "Code",
  "Item Name",
  "Before",
  "After",
  "Diff",
  "%",
  "Unit",
];
// Menambahkan data
const data = [
  [
    "",
    "ADJ0257",
    "01/12/2023",
    "Bar",
    "Confirmed",
    10305,
    "Fresh Lemon",
    89,
    3,
    -1,
    -84,
    "Kg",
  ],
  [
    "",
    "ADJ0257",
    "01/12/2023",
    "Bar",
    "Confirmed",
    10305,
    "Fresh Lemon",
    89,
    3,
    -1,
    -84,
    "pcs",
  ],
];

const footter = ["", "", "", "", "", "", "TOTAL", "", "", "", "", ""];

// menambahkan data ke worksheet
const header1Row = worksheet.addRow(headers);
const header2Row = worksheet.addRow(headers2);
data.forEach((row, index) => {
  const dataRow = worksheet.addRow(row);

  // Memberikan border pada seluruh data row
  dataRow.eachCell((cell, cellNumber) => {
    cell.border = {
      top: { style: index === 0 ? "thin" : "none" }, // Add top border for the first data row
      right: { style: cellNumber === dataRow.cellCount ? "thin" : "none" }, // Add right border for the last cell in each row
    };
  });
});
const footerRow = worksheet.addRow(footter);

// style

header1Row.eachCell((cell) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "B0B0B0" },
  };
  cell.border = {
    top: { style: "thin" },
    bottom: { style: "thin" },
  };
});
header2Row.eachCell((cell) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "B0B0B0" },
  };
  cell.border = {
    top: { style: "thin" },
    bottom: { style: "thin" },
  };
});
footerRow.eachCell((cell, cellNumber) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "B0B0B0" },
  };
  cell.border = {
    top: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: cellNumber === footerRow.cellCount ? "thin" : "none" },
  };
});

worksheet.getCell("L5").border = {
  right: { style: "thin" },
  top: { style: "thin" },
  bottom: { style: "thin" },
};
worksheet.getCell("L6").border = {
  right: { style: "thin" },
  top: { style: "thin" },
  bottom: { style: "thin" },
};

worksheet.views = [{ state: "frozen", xSplit: 0, ySplit: 6, activeCell: "A6" }];

worksheet.getColumn(6).width = 6;
worksheet.getColumn(7).width = 30;
worksheet.getColumn(8).width = 8;
worksheet.getColumn(9).width = 8;
worksheet.getColumn(10).width = 8;
worksheet.getColumn(11).width = 6;
worksheet.getColumn(12).width = 6;

worksheet.getColumn(12).alignment = {
  vertical: "middle",
  horizontal: "right",
};
worksheet.getCell("H6").alignment = {
  vertical: "middle",
  horizontal: "right",
};
worksheet.getCell("I6").alignment = {
  vertical: "middle",
  horizontal: "right",
};
worksheet.getCell("J6").alignment = {
  vertical: "middle",
  horizontal: "right",
};

// Menyimpan Workbook ke File Excel
workbook.xlsx
  .writeFile("output.xlsx")
  .then(() => {
    console.log("File Excel berhasil disimpan.");
  })
  .catch((error) => {
    console.error("Gagal menyimpan file Excel:", error);
  });
