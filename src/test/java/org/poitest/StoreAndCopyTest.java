package org.poitest;

import org.junit.Before;
import org.junit.After;
import org.junit.Test;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import static junit.framework.Assert.assertNotNull;

/**
 * Date: Apr 15, 2009
 *
 * @author Leonid Vysochyn
 */
public class StoreAndCopyTest {
    public static final String TEMPLATE_FILE_NAME = "storecopy.xls";
    public static final String MODIFIED_FILE_NAME = "target/storecopy_modified.xls";

    InputStream templateInputStream;
    Workbook workbook;
    
    @Before
    public void createWorkbookFromFile() throws Exception {
        InputStream resourceAsStream = getClass().getResourceAsStream(TEMPLATE_FILE_NAME);
        assertNotNull("Can not open input stream for file name = " + TEMPLATE_FILE_NAME, resourceAsStream);
        templateInputStream = new BufferedInputStream(resourceAsStream);
        workbook = WorkbookFactory.create( templateInputStream );
        assertNotNull("Workbook must be not null", workbook);
    }

    @After
    public void closeTemplateInputStream() throws IOException {
        templateInputStream.close();
    }

    @Test
    public void storeAndCopyRows() throws Exception {
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        Sheet sheet = workbook.getSheetAt(0);
        XlsCellInfo cellOneInfo = Util.readCell(sheet, 1, 0);
        XlsCellInfo cellTwoInfo = Util.readCell(sheet, 1, 1);
        copyRows(sheet, 0, 1);
        Row newRow = sheet.getRow(2);
        if( newRow == null ){
            newRow = sheet.createRow(2);
        }
        Util.writeCellToRow(newRow, cellOneInfo);
        Util.writeCellToRow(newRow, cellTwoInfo);
        workbook.write( resultStream );
        resultStream.close();
        templateInputStream.close();
        assertRowsInInputOutputExcelFiles("Sheet1", 1, 2);
    }

    private void assertRowsInInputOutputExcelFiles(String sheetName, int srcRowNum, int destRowNum) throws Exception {
        InputStream srcInputStream = getClass().getResourceAsStream(TEMPLATE_FILE_NAME);
        InputStream destInputStream = new FileInputStream(MODIFIED_FILE_NAME);
        Workbook srcWorkbook = WorkbookFactory.create( srcInputStream );
        Workbook destWorkbook = WorkbookFactory.create( destInputStream );
        Sheet srcSheet = srcWorkbook.getSheet(sheetName);
        Sheet destSheet = destWorkbook.getSheet(sheetName);
        Row srcRow = srcSheet.getRow( srcRowNum );
        Row destRow = destSheet.getRow( destRowNum );
        assertRowCellsEqual(srcRow, destRow);
    }

    private void assertRowCellsEqual(Row srcRow, Row destRow) {
        for (Cell srcCell : srcRow) {
            int colIndex = srcCell.getColumnIndex();
            Cell destCell = destRow.getCell(colIndex);
            Util.assertCellsEquality(srcCell, destCell);
        }
    }

    private void copyRows(Sheet sheet, int srcRowNum, int destRowNum) {
        Row srcRow = sheet.getRow(srcRowNum);
        Row destRow = sheet.getRow(destRowNum);
        if( destRow == null ){
            destRow = sheet.createRow( destRowNum );
        }
        copyRows( srcRow, destRow );
    }

    private void copyRows(Row srcRow, Row destRow) {
        for (Cell srcCell : srcRow) {
            int colIndex = srcCell.getColumnIndex();
            Cell destCell = destRow.getCell(colIndex);
            if (destRow.getCell(colIndex) == null) {
                destCell = destRow.createCell(colIndex);
            }
            copyCells(srcCell, destCell);
        }
    }

    private void copyCells(Cell srcCell, Cell destCell) {
        Util.copyCells(srcCell, destCell);
    }


}
