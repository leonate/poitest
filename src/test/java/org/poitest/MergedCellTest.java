package org.poitest;

import junit.framework.Assert;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.*;

import static junit.framework.Assert.assertNotNull;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12 1:05 PM
 */
public class MergedCellTest {
    public static final String TEMPLATE_FILE_NAME = "mergedcells.xlsx";
    public static final String MODIFIED_FILE_NAME = "target/mergedcells_modified.xlsx";
    public static final String MODIFIED_FILE_NAME2 = "target/mergedcells_modified2.xlsx";
    public static final String MODIFIED_FILE_NAME3 = "target/mergedcells_modified3.xlsx";

    InputStream templateInputStream;
    Workbook workbook;

    @Before
    public void createWorkbookFromFile() throws Exception {
        InputStream resourceAsStream = getClass().getResourceAsStream(TEMPLATE_FILE_NAME);
        assertNotNull("Can not open input stream for file name = " + TEMPLATE_FILE_NAME, resourceAsStream);
        templateInputStream = new BufferedInputStream(resourceAsStream);
        workbook = WorkbookFactory.create(templateInputStream);
        assertNotNull("Workbook must be not null", workbook);
    }

    @After
    public void closeTemplateInputStream() throws IOException {
        templateInputStream.close();
    }

    @Test
    public void checkMergedRegions(){
        Sheet sheet = workbook.getSheetAt(0);
        Assert.assertEquals("Incorrect number of merged regions", 1, sheet.getNumMergedRegions());
        Assert.assertEquals("Incorrect merged region", new CellRangeAddress(0,1,0,1).toString(),sheet.getMergedRegion(0).toString());
    }

    @Test
    public void readMergedCell(){
        Sheet sheet = workbook.getSheetAt(0);
        assertEquals("Merged cell value is read incorrectly", "Merged Cell", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Merged cell part should be blank", Cell.CELL_TYPE_BLANK, sheet.getRow(0).getCell(1).getCellType());
        assertEquals("Merged cell part should be blank", Cell.CELL_TYPE_BLANK, sheet.getRow(1).getCell(0).getCellType());
        assertEquals("Merged cell part should be blank", Cell.CELL_TYPE_BLANK, sheet.getRow(1).getCell(1).getCellType());
    }
    
    @Test
    public void readAdjacentCellsToMergedCell(){
        Sheet sheet = workbook.getSheetAt(0);
        assertEquals("Adjacent to merged cell value is read incorrectly", 1, sheet.getRow(0).getCell(2).getNumericCellValue(), 1e-6);
        assertEquals("Adjacent to merged cell value is read incorrectly", 2, sheet.getRow(1).getCell(2).getNumericCellValue(), 1e-6);
        assertEquals("Adjacent to merged cell value is read incorrectly", 3, sheet.getRow(2).getCell(0).getNumericCellValue(), 1e-6);
        assertEquals("Adjacent to merged cell value is read incorrectly", 4, sheet.getRow(2).getCell(1).getNumericCellValue(), 1e-6);
        assertEquals("Adjacent to merged cell value is read incorrectly", 5, sheet.getRow(2).getCell(2).getNumericCellValue(), 1e-6);
    }

    @Test
    public void copyMergedCell() throws IOException {
        Sheet sheet = workbook.getSheetAt(0);
        Cell srcCell = sheet.getRow(0).getCell(0);
        Cell destCell = sheet.createRow(10).createCell(4);
        Util.copyCells(srcCell, destCell);
        sheet.addMergedRegion(new CellRangeAddress(10, 11, 4, 5));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 2));
        Assert.assertEquals("Incorrect number of merged regions", 3, sheet.getNumMergedRegions());
        Assert.assertEquals("Incorrect merged region", new CellRangeAddress(10, 11, 4, 5).toString(), sheet.getMergedRegion(1).toString());
        Assert.assertEquals("Incorrect merged region", new CellRangeAddress(2,2,1,2).toString(),sheet.getMergedRegion(2).toString());
        Assert.assertEquals(4, sheet.getRow(2).getCell(1).getNumericCellValue(), 1e-6);
        Assert.assertEquals(5, sheet.getRow(2).getCell(2).getNumericCellValue(), 1e-6);
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        workbook.write( resultStream );
        resultStream.close();
    }
    
    @Test
    public void removeMergedRegion() throws IOException {
        Sheet sheet = workbook.getSheetAt(0);
        Assert.assertEquals(1, sheet.getNumMergedRegions());
        sheet.getRow(0).getCell(1).setCellValue("test value");
        sheet.removeMergedRegion(0);
        Assert.assertEquals(0, sheet.getNumMergedRegions());
        OutputStream resultStream = new FileOutputStream( MODIFIED_FILE_NAME2 );
        workbook.write( resultStream );
        resultStream.close();
    }

    @Test
    public void nestedMergedRegions() throws IOException {
        Sheet sheet = workbook.getSheetAt(0);
        Assert.assertEquals(1, sheet.getNumMergedRegions());
        sheet.removeMergedRegion(0);
        sheet.addMergedRegion(new CellRangeAddress(0,1,0,2));
        Assert.assertEquals(1, sheet.getNumMergedRegions());
        OutputStream resultStream = new FileOutputStream( MODIFIED_FILE_NAME3 );
        workbook.write(resultStream);
        resultStream.close();
    }
    
}
