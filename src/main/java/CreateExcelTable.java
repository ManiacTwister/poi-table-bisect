
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

class CreateExcelTable {

    public static void main(String[] args) throws Exception {

        try ( XSSFWorkbook workbook = new XSSFWorkbook();  FileOutputStream fileout = new FileOutputStream("/home/robin/Dokumente/share/5.1.0.xlsx")) {

            //prepairing the sheet
            XSSFSheet sheet = workbook.createSheet();

            String[] tableHeadings = new String[]{"Heading1", "Heading2", "Heading3", "Heading4"};
            String tableName = "Table1";
            int firstRow = 0; //start table in row 1
            int firstCol = 0; //start table in column A
            int rows = 6; //we have to populate headings row, 4 data rows and 1 totals row
            int cols = 4; //three columns in each row

            for (int r = 0; r < rows; r++) {
                XSSFRow row = sheet.createRow(firstRow + r);
                for (int c = 0; c < cols; c++) {
                    XSSFCell localXSSFCell = row.createCell(firstCol + c);
                    if (r == 0) {
                        localXSSFCell.setCellValue(tableHeadings[c]);
                    } else if (r == 5) {
                        //totals row content will be set later
                    } else {
                        localXSSFCell.setCellValue(r + c);
                    }
                }
            }

            //create the table
            CellReference topLeft = new CellReference(sheet.getRow(firstRow).getCell(firstCol));
            CellReference bottomRight = new CellReference(sheet.getRow(firstRow + rows - 1).getCell(firstCol + cols - 1));
            AreaReference tableArea = workbook.getCreationHelper().createAreaReference(topLeft, bottomRight);
            XSSFTable dataTable = sheet.createTable(tableArea);
            dataTable.setName(tableName);
            dataTable.setDisplayName(tableName);

            //this styles the table as Excel would do per default
            dataTable.getCTTable().addNewTableStyleInfo();
            XSSFTableStyleInfo style = (XSSFTableStyleInfo) dataTable.getStyle();
            style.setName("TableStyleMedium9");
            style.setShowColumnStripes(false);
            style.setShowRowStripes(true);
            style.setFirstColumn(false);
            style.setLastColumn(false);

            //this sets auto filters
            dataTable.getCTTable().addNewAutoFilter().setRef(tableArea.formatAsString());

            //this sets totals properties to table and totals formulas to sheet
            XSSFRow totalsRow = dataTable.getXSSFSheet().getRow(tableArea.getLastCell().getRow());
            for (int c = 0; c < dataTable.getCTTable().getTableColumns().getTableColumnList().size(); c++) {
                switch (c) {
                    case 0:
                        dataTable.getCTTable().getTableColumns().getTableColumnList().get(c).setTotalsRowLabel("Totals: ");
                        totalsRow.getCell(tableArea.getFirstCell().getCol() + c).setCellValue("Totals: ");
                        break;
                    case 1:
                        dataTable.getCTTable().getTableColumns().getTableColumnList().get(c).setTotalsRowFunction(org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction.SUM);
                        totalsRow.getCell(tableArea.getFirstCell().getCol() + c).setCellFormula("SUBTOTAL(109," + tableName + "[" + tableHeadings[c] + "])");
                        break;
                    case 2:
                        dataTable.getCTTable().getTableColumns().getTableColumnList().get(c).setTotalsRowFunction(org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction.SUM);
                        totalsRow.getCell(tableArea.getFirstCell().getCol() + c).setCellFormula("SUBTOTAL(109," + tableName + "[" + tableHeadings[c] + "])");
                        break;
                    case 3:
                        dataTable.getCTTable().getTableColumns().getTableColumnList().get(c).setTotalsRowLabel("Totals: ");
                        totalsRow.getCell(tableArea.getFirstCell().getCol() + c).setCellValue("Totals: ");
                        break;
                    default:
                        break;
                }
            }
            //this shows the totals row
            dataTable.getCTTable().setTotalsRowCount(1);

            workbook.write(fileout);
        }

    }
    }