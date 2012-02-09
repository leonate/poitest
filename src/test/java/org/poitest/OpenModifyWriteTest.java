package org.poitest;

import org.junit.Test;
import org.junit.Before;
import org.junit.After;
import org.junit.Ignore;
import static org.junit.Assert.assertEquals;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import static junit.framework.Assert.assertNotNull;

/**
 * Date: Mar 11, 2009
 *
 * @author Leonid Vysochyn
 */
public class OpenModifyWriteTest {
    public static final String TEMPLATE_FILE_NAME = "reading.xls";
    public static final String MODIFIED_FILE_NAME = "target/modified.xls";

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
    public void modifyStringValue() throws Exception {
        int sheetNum = 0;
        int rowNum = 0;
        int cellNum = 0;
        String cellValue = "A test";
        writeCellValue(sheetNum, rowNum, cellNum, cellValue);
        assertModifiedCellValue(sheetNum, rowNum, cellNum, cellValue);
    }

    private void writeCellValue(int sheetNum, int rowNum, int cellNum, String cellValue) throws IOException {
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(cellNum);
        cell.setCellValue(cellValue);
        workbook.write( resultStream );
        resultStream.flush();
        resultStream.close();
    }

    private static void assertModifiedCellValue(int sheetNum, int rowNum, int cellNum, String cellValue) throws Exception {
        InputStream inputStream = new FileInputStream(MODIFIED_FILE_NAME);
        Workbook reopenedWorkbook = WorkbookFactory.create( inputStream );
        Sheet sheet = reopenedWorkbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(cellNum);
        assertEquals("Written and read String values are different", cellValue, cell.getStringCellValue());
        inputStream.close();
    }

    @Test
    public void createRowOverExistingOne() throws Exception {
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.createRow(1);
        assertEquals("First Cell num is incorrect", -1, row.getFirstCellNum());
        Cell cell = row.createCell(2);
        cell.setCellValue("New Cell");
        workbook.write( resultStream );
        resultStream.close();
        assertEquals("First Cell num is incorrect", 2, row.getFirstCellNum());
        assertRowCount(0, 4);
        assertModifiedCellValue(0, 1, 2, "New Cell");
    }


    @Ignore
    public void shiftRows() throws Exception {
        int sheetNum = 0;
        int startRow = 1;
        int endRow = 2;
        int shiftNum = 2;
        shiftRows(sheetNum, startRow, endRow, shiftNum);
        assertRowCount(sheetNum, 6);
    }

    private void shiftRows(int sheetNum, int startRow, int endRow, int shiftNum) throws IOException {
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        Sheet sheet = workbook.getSheetAt(sheetNum);
        sheet.shiftRows(startRow, endRow, shiftNum);
        workbook.write( resultStream );
        resultStream.close();
    }

    private void assertRowCount(int sheetNum, int expectedRowCount) throws Exception {
        InputStream inputStream = new FileInputStream(MODIFIED_FILE_NAME);
        Workbook reopenedWorkbook = WorkbookFactory.create( inputStream );
        Sheet sheet = reopenedWorkbook.getSheetAt(sheetNum);
        assertEquals( "Incorrect number of rows", expectedRowCount, sheet.getLastRowNum() + 1);
    }

}
