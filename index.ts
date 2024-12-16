import ExcelJS from "exceljs";
const test = [
  {
    id: 1,
    name: "Bold Text",
    style: { font: { bold: true } },
  },
  {
    id: 2,
    name: "Italic Text",
    style: { font: { italic: true } },
  },
  {
    id: 3,
    name: "Underlined Text",
    style: { font: { underline: true } },
  },
  {
    id: 4,
    name: "Red Text",
    style: { font: { color: { argb: "FFFF0000" } } },
  },
  {
    id: 5,
    name: "Large Text",
    style: { font: { size: 16 } },
  },
  {
    id: 6,
    name: "Yellow Fill",
    style: {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      },
    },
  },
  {
    id: 7,
    name: "Thick Border",
    style: {
      border: {
        top: { style: "thick" },
        left: { style: "thick" },
        bottom: { style: "thick" },
        right: { style: "thick" },
      },
    },
  },
  {
    id: 8,
    name: "Centered Text",
    style: {
      alignment: { horizontal: "center", vertical: "middle" },
    },
  },
  {
    id: 9,
    name: "Date Format",
    value: new Date(2023, 0, 1),
    style: { numFmt: "dd/mm/yyyy" },
  },
  {
    id: 10,
    name: "Currency Format",
    value: 1234.56,
    style: { numFmt: '"$"#,##0.00' },
  },
  {
    id: 11,
    name: "Formula Example",
    formula: "SUM(C2:C10)",
  }, {
    id: 12,
    name: "Formula Example",
    merge: "A6:C6", // Intervalo de células que serão mescladas
    style: {
      font: { bold: true, size: 14 },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD3D3D3" }, // Cor de preenchimento cinza claro
      },
    },

  }
];



async function generateExcel(data: any) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("AppSetting");

  // Configurar colunas
  sheet.columns = [
    { header: "ID", key: "id", width: 10 },
    { header: "Name", key: "name", width: 30 },
    { header: "Value", key: "value", width: 20 },
  ];

  // Adicionar linhas
  data.forEach((item: any) => {
    const row = sheet.addRow({
      id: item.id,
      name: item.name,
      value: item.value,
    });

    // Verificar se é para mesclar células
    if (item.merge) {
      sheet.mergeCells(item.merge);
      const mergedCell = sheet.getCell(item.merge.split(":")[0]); // Obter a primeira célula do intervalo
      mergedCell.value = item.name; // Preencher com o valor do campo `name`

      // Aplicar estilos às células mescladas
      if (item.style) {
        if (item.style.font) mergedCell.font = item.style.font;
        if (item.style.fill) mergedCell.fill = item.style.fill;
        if (item.style.alignment) mergedCell.alignment = item.style.alignment;
      }
    } else {
      // Aplicar estilos dinamicamente às células não mescladas
      if (item.style) {
        row.eachCell((cell) => {
          if (item.style.font) cell.font = item.style.font;
          if (item.style.fill) cell.fill = item.style.fill;
          if (item.style.border) cell.border = item.style.border;
          if (item.style.alignment) cell.alignment = item.style.alignment;
          if (item.style.numFmt) cell.numFmt = item.style.numFmt;
        });
      }

      // Adicionar fórmula dinamicamente
      if (item.formula) {
        const cell = row.getCell(3); // Coluna "Value"
        cell.value = { formula: item.formula, result: 0 }; // Valor inicial
      }
    }
  });

  // Salvar arquivo
  await workbook.xlsx.writeFile("./test.xlsx");
}

// Executar função dinâmica
generateExcel(test).catch(console.error);
