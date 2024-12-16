import ExcelJS, { Cell } from "exceljs";

const test = [
  {
    id: 1,
    name: "test",
    style: {
      font: { bold: true },
      fill: {
        type: "pattern" as const, // Declara o tipo literal explicitamente
        pattern: "solid" as const,
        fgColor: { argb: "FFFFE599" },
      },
    },
  },
  { id: 2, name: "test" },
  { id: 3, name: "test" },
  { id: 4, name: "test" },
  { id: 5, name: "test" },
  { id: 6, name: "test" },
];

async function generateExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("StyledSheet");

  // Configurando as colunas
  sheet.columns = [
    { header: "ID", key: "id", width: 10 },
    { header: "Name", key: "name", width: 20 },
  ];

  // Adicionando linhas e aplicando estilos
  test.forEach((item) => {
    const row = sheet.addRow({ id: item.id, name: item.name });
    if (item.style) {
      row.eachCell((cell: Cell) => {
        if ("font" in item.style) {
          cell.font = item.style.font;
        }
        if ("fill" in item.style) {
          cell.fill = item.style.fill;
        }
      });
    }
  });

  // Salvando o arquivo Excel
  await workbook.xlsx.writeFile("./test.xlsx");
  console.log("Excel file created: test.xlsx");
}

generateExcel().catch(console.error);
