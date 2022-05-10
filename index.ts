import Excel from "exceljs";

(async () => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("spreadsheet.xlsx");
  const seenStyles = new Set();
  workbook.eachSheet((sheet) => {
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (seenStyles.has(cell.style)) {
          console.log(
            `${cell.$col$row} style object was already seen on another cell`
          );
        }
        seenStyles.add(cell.style);
      });
    });
  });
})();
