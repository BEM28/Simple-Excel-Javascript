const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet1");

// Set values for merged cells and align them
const mergeCellOptionsCenter = {
  horizontal: "center",
  vertical: "middle",
};
const mergeCellOptionsLeft = {
  horizontal: "left",
  vertical: "middle",
};

worksheet.getCell("A1").value = "Laporan Detail Menu Pemakaian POS";
worksheet.getCell("A1").font = { size: 16 };
worksheet.getCell("A1").alignment = mergeCellOptionsCenter;

worksheet.getCell("A2").value = "Term : 01-December-2023 - 31-December-2023";
worksheet.getCell("A2").alignment = mergeCellOptionsCenter;

worksheet.getCell("A3").value = "Category : All";
worksheet.getCell("A3").alignment = mergeCellOptionsCenter;

worksheet.getCell("A4").value = "Category : All";
worksheet.getCell("A4").alignment = mergeCellOptionsCenter;

worksheet.addRow([]);
worksheet.addRow([]);

// Menambahkan header
const headers = ["No", "Item Code", "Item Name", "Qty", ""];
// Menambahkan data
const data = [
  [1, 10101, "Samcan", 22.62, "Kg"],
  [2, 10102, "Sasa", 22.64, "Pcs"],
];

// menambahkan data ke worksheet
worksheet.mergeCells("A1:E1");
worksheet.mergeCells("A2:E2");
worksheet.mergeCells("A3:E3");
worksheet.mergeCells("A4:E4");

const header1Row = worksheet.addRow(headers);

header1Row.eachCell((cell, index) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "B0B0B0" },
  };
  if (index == 5) {
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
  cell.font = {
    bold: true,
  };
});

// Set color for the data rows
data.forEach((rowData) => {
  const dataRow = worksheet.addRow(rowData);
  dataRow.eachCell((cell, index) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: index <= 3 ? "feffcd" : "faf904" },
    };
    cell.border = {
      top: { style: index === 0 ? "thin" : "none" },
      bottom: { style: "thin" },
      right: { style: index === dataRow.cellCount ? "thin" : "none" },
    };
    if (index == 1 || index == 5) {
      cell.alignment = mergeCellOptionsCenter;
    } else {
      cell.alignment = mergeCellOptionsLeft;
    }
  });
});

worksheet.getColumn(1).width = 4;
worksheet.getColumn(2).width = 12;
worksheet.getColumn(3).width = 34;
worksheet.getColumn(4).width = 6;
worksheet.getColumn(5).width = 5;

// Menyimpan Workbook ke File Excel
workbook.xlsx
  .writeFile("output1.xlsx")
  .then(() => {
    console.log("File Excel berhasil disimpan.");
  })
  .catch((error) => {
    console.error("Gagal menyimpan file Excel:", error);
  });
