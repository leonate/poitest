package org.poitest;

import org.apache.poi.ss.usermodel.*;
import org.junit.*;
import static org.junit.Assert.assertEquals;

import java.io.*;
import java.util.Date;
import java.text.SimpleDateFormat;

import static junit.framework.Assert.assertNotNull;

/**
 * Date: Mar 11, 2009
 *
 * @author Leonid Vysochyn
 */
public class OpenReadTest {
    public static final String TEMPLATE_FILE_NAME = "reading.xls";

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
    public void readStringValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        assertEquals("String cell value is read incorrectly", "Employees", cell.getStringCellValue());
    }

    @Test
    public void readIntValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        assertEquals("Int cell value is read incorrectly", 12345, (int)cell.getNumericCellValue());
    }

    @Test
    public void readDoubleValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(1);
        assertEquals("Int cell value is read incorrectly", 12.345, cell.getNumericCellValue(), 1e-6);
    }

    @Test
    public void readCurrencyValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(2);
        assertEquals("Currency cell value is read incorrectly", 1234.102, cell.getNumericCellValue(), 1e-6);
    }

    @Test
    public void readDateValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(3);
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date calendarDate = dateFormat.parse("03/14/2009");
        assertEquals("Date cell value is read incorrectly", calendarDate,
                cell.getDateCellValue());
    }

    @Test
    public void readPercentageValue() throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(4);
        assertEquals("Percentage cell value is read incorrectly", 0.2, cell.getNumericCellValue(), 1e-6);
    }

    @Test
    public void readCellComment(){
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(2);
        Cell cell = row.getCell(0);
        assertEquals("Cell comment is incorrect", "Total row", cell.getCellComment().getString().getString());
    }

}
