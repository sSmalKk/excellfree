import ExcelJS, { Cell } from "exceljs";

const test = [
  {
    id: 1,
    name: "Bold Text",
    style: {
      font: { bold: true },
    },
  },
  {
    id: 2,
    name: "Italic Text",
    style: {
      font: { italic: true },
    },
  },
  {
    id: 3,
    name: "Underlined Text",
    style: {
      font: { underline: true },
    },
  },
  {
    id: 4,
    name: "Red Text",
    style: {
      font: { color: { argb: "FFFF0000" } },
    },
  },
  {
    id: 5,
    name: "Large Text",
    style: {
      font: { size: 16 },
    },
  },
  {
    id: 6,
    name: "Yellow Fill",
    style: {
      fill: {
        type: "pattern" as const,
        pattern: "solid" as const,
        fgColor: { argb: "FFFFFF00" },
      },
    },
  },
  {
    id: 7,
    name: "Thick Border",
    style: {
      border: {
        top: { style: "thick" as const },
        left: { style: "thick" as const },
        bottom: { style: "thick" as const },
        right: { style: "thick" as const },
      },
    },
  },
  {
    id: 8,
    name: "Centered Text",
    style: {
      alignment: {
        horizontal: "center" as const,
        vertical: "middle" as const,
      },
    },
  },
  {
    id: 9,
    name: "Date Format",
    value: new Date(2023, 0, 1),
    style: {
      numFmt: "dd/mm/yyyy",
    },
  },
  {
    id: 10,
    name: "Currency Format",
    value: 1234.56,
    style: {
      numFmt: '"$"#,##0.00',
    },
  },
  {
    id: 11,
    name: "Formula Example",
    formula: "SUM(A2:A10)",
  },
];

async function generateExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("StyledSheet");

  // Configurando as colunas
  sheet.columns = [
    { header: "ID", key: "id", width: 10 },
    { header: "Name", key: "name", width: 30 },
    { header: "Value/Formula", key: "value", width: 20 },
  ];

  // Adicionando linhas e aplicando estilos
  test.forEach((item) => {
    const row = sheet.addRow({ id: item.id, name: item.name, value: item.value });
    if (item.style) {
      row.eachCell((cell: Cell) => {
        if (item.style.font) {
          cell.font = item.style.font;
        }
        if (item.style.fill) {
          cell.fill = item.style.fill;
        }
        if (item.style.border) {
          cell.border = item.style.border;
        }
        if (item.style.alignment) {
          cell.alignment = item.style.alignment;
        }
        if (item.style.numFmt) {
          cell.numFmt = item.style.numFmt;
        }
      });
    }
    if (item.formula) {
      row.getCell(3).value = { formula: item.formula, result: 0 }; // Use "0" como valor padrão
    }
  });

  // Mesclando células (Exemplo)
  sheet.mergeCells("A13:C13");
  const mergedCell = sheet.getCell("A13");
  mergedCell.value = "Merged Cells Example";
  mergedCell.alignment = { horizontal: "center", vertical: "middle" };
  mergedCell.font = { bold: true, size: 14 };

  // Salvando o arquivo Excel
  await workbook.xlsx.writeFile("./test.xlsx");
  console.log("Excel file created: test.xlsx");
}

generateExcel().catch(console.error);
