import ExcelJS from "exceljs";

const test = [
  { id: 1, name: "John" },
  { id: 2, name: "Alice" },
  { id: 3, name: "Bob" },
];

async function generateExcel() {
  const excel = new ExcelJS.Workbook();
  const sheet = excel.addWorksheet("App");

  sheet.columns = Object.keys(test[0]).map(key => ({
    header: key,
    key: key,
    width: 25,
  }));

  sheet.addRows(test);

  test.map((item, rowIndex) => {
    const row = sheet.getRow(rowIndex + 2);

    // Se a chave for 'id', aplica fonte em negrito
    if ('id' in item) {
      const fontStyle: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FF5733" } };
      row.eachCell((cell, colNumber) => {
        if (colNumber === 1) { // Coluna 'id'
          cell.style.font = fontStyle;
        }
      });
    }

    // Se a chave for 'name', aplica fundo amarelo
    if ('name' in item) {
      const fillStyle: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } };
      row.eachCell((cell, colNumber) => {
        if (colNumber === 2) { // Coluna 'name'
          cell.style.fill = fillStyle;
        }
      });
    }
  });

  await excel.xlsx.writeFile("./test.xlsx");
}

generateExcel();
