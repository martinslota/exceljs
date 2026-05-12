const ExcelJS = verquire('exceljs');

// Drawings whose root element isn't `xdr:wsDr` (e.g. `<c:userShapes>` chart
// overlays, or `<wsDr>` using the default namespace instead of the `xdr:`
// prefix) used to parse into a `null` model. That null then crashed
// `XLSX.reconcile` and `WorksheetXform.reconcile` when the worksheet still
// referenced the drawing part. Both files below are real-world examples
// produced by other tools.
describe('github issues', () => {
  it('tolerates a chart user shapes overlay drawing', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-chart-user-shapes.xlsx')
      .then(() => {
        expect(wb.worksheets.length).to.be.greaterThan(0);
        const [sheet] = wb.worksheets;
        expect(sheet.rowCount).to.be.greaterThan(0);
      });
  });

  it('tolerates a drawing that uses the default namespace instead of xdr:', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-drawing-default-namespace.xlsx')
      .then(() => {
        expect(wb.worksheets.length).to.be.greaterThan(0);
        const [sheet] = wb.worksheets;
        expect(sheet.rowCount).to.be.greaterThan(0);
      });
  });
});
