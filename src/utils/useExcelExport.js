import ExcelJS  from "exceljs";
import { saveAs } from "file-saver";



const defaultOptions = {
  name: "Export",
  data: [
    [new Date("2019-07-20"), 70.1],
    [new Date("2019-07-21"), 70.6],
    [new Date("2019-07-22"), 70.1]
  ],
  tableOptions: {
    name: "TableExport",
    ref: "A1",
    style: {
      theme: "TableStyleLight3",
      showRosheettripes: true
    },
    columns: [
      { name: "ID", filterButton: true, width: 10, },
      { name: "FirstName", filterButton: false, width: 32 },
      { name: "LastName", filterButton: true },
      { name: "Prefix", filterButton: false },
      { name: "Position", filterButton: true },
      { name: "Picture", filterButton: false },
      { name: "BirthDate", filterButton: true },
      { name: "HireDate", filterButton: false },
      { name: "Notes", filterButton: true },
      { name: "Adress", filterButton: false },
      { name: "State", filterButton: false },
      { name: "City", filterButton: true },
      { name: "SaleAmount", filterButton: false },
    ]
  }
};

export const useExcelExport = () => {
  const generate = (data, workbookOptions, tableOptions) => {

    const length = data.length + 1;

    const configuration = { ...defaultOptions.tableOptions, ...tableOptions };

    console.log("CONFIG", configuration);

    var workbook = new ExcelJS.Workbook();
    var sheet = workbook.addWorksheet(workbookOptions.name);
    

    var column = sheet.getColumn(2);
    column.outlineLevel = 1;
    
    // console.log(data);

    sheet.addTable({
      name: configuration.name.replace(' ', '_'),
      ref: configuration.ref,
      headerRow: configuration.headerRow,
      totalsRow: configuration.totalsRow,
      style: configuration.style,
      columns: configuration.columns,
      rows: data
    });

    // define columns
    sheet.columns = [
      { header: 'ID', key: 'id', width: 6,},
      { header: 'FirstName', key: 'firstname', width: 12,},
      { header: 'LastName', key: 'lastname', width: 12, },
      { header: 'Prefix', key: 'Prefix', width: 8 },
      { header: 'Position', key: 'Position', width: 18 },
      { header: 'Picture', key: 'Picture', width: 26,},
      { header: 'BirthDate', key: 'BirthDate', width: 14, style: { numFmt: 'dd/mm/yyyy' } },
      { header: 'HireDate', key: 'HireDate', width: 14, style: { numFmt: 'dd/mm/yyyy' } },
      { header: 'Notes', key: 'Notes', width: 32 },
      { header: 'Adress', key: 'Adress', width: 18 },
      { header: 'State', key: 'State', width: 12 },
      { header: 'City', key: 'City', width: 12 },
      { header: 'SaleAmount', key: 'SaleAmount', width: 10 },
    ];

    // get Row font
    sheet.getRow(1).font = {
      name: 'Arial',
      family: 4,
      size: 10,
      underline: false,
      bold: true,
      color: {argb: '1419F8'},
    };
    for (var x = 1; x <= length; x ++ ) {
      sheet.getRow(x).alignment = { vertical: 'middle', horizontal: 'center' };
    }
  
    // Get Cell
    for(var i = 2; i <= length; i ++) {
      sheet.getCell(`A[${i}]`).font = {
        name: 'Arial',
        family: 4,
        size: 10,
        underline: false,
        bold: true,
        color: {argb: '14190F'},
      }
    }

    // Border
    sheet.getCell('B1').border = {
      top: {style:'double', color: {argb:'FF00FF00'}},
      left: {style:'double', color: {argb:'FF00FF00'}},
      bottom: {style:'double', color: {argb:'FF00FF00'}},
      right: {style:'double', color: {argb:'FF00FF00'}}
    };

    // Fill
    sheet.getColumn('Position').eachCell(function(cell) {
      if(cell.value === 'CEO') {
        console.log(cell.value)
        cell.fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FFFF0000'}
        }
      }
    })

    sheet.getColumn('Adress').eachCell(function(cell) {
      if(cell.value) {
        cell.fill = {
          type: 'gradient',
          gradient: 'angle',
          degree: 0,
          stops: [
            {position:0, color:{argb:'FF0000FF'}},
            {position:0.5, color:{argb:'FFFFFFFF'}},
            {position:1, color:{argb:'FF0000FF'}}
          ]
        }
      }
    })
    
    workbook.xlsx.writeBuffer().then(data => {
      const blob = new Blob([data], { type: "text/plain;charset=utf-8" });
      saveAs(blob, `${workbookOptions.name}.xlsx`);
    });
    
  };

  return {
    generate
  };
};
