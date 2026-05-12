const ExcelJS = verquire('exceljs');

// Comments parts stored at non-canonical paths (e.g. `xl/sheet1_comments.xml`
// instead of `xl/commentsN.xml`) are not picked up by the loader's regex.
// When the worksheet's rels still reference such a target, reconcile used to
// crash dereferencing `options.comments[rel.Target].comments`. We now skip
// the dangling rel so the workbook reads without error.
describe('github issues', () => {
  it('tolerates a worksheet rel that points at a non-standard comments part', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-non-standard-comments.xlsx')
      .then(() => {
        expect(wb.worksheets.length).to.be.greaterThan(0);
        const [sheet] = wb.worksheets;
        expect(sheet.rowCount).to.be.greaterThan(0);
      });
  });
});
