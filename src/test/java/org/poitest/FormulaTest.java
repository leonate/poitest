package org.poitest;

import org.apache.poi.ss.usermodel.*;
import org.junit.Before;
import org.junit.After;
import org.junit.Test;
import static org.junit.Assert.assertEquals;

import java.io.*;

import static junit.framework.Assert.assertNotNull;

/**
 * Date: Mar 12, 2009
 *
 * @author Leonid Vysochyn
 */
public class FormulaTest {

    public static final String TEMPLATE_FILE_NAME = "formulas.xls";
    public static final String MODIFIED_FILE_NAME = "target/formulas_modified.xls";

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
    public void readSimpleFormula(){
        assertFormula(0, 3, 1, "sum(b1:b3)", "Simple formula can not be read correctly");
    }

    @Test
    public void readIncompleteFormula(){
        assertFormula(0, 5, 1, "B4+XValue", "First Incomplete formula  can not be read correctly");
        assertFormula(0, 5, 2, "AVERAGE(B1,avalue)", "Second Incomplete formula  can not be read correctly");
    }

    @Test
    public void writeFormula() throws Exception {
        String formula = "average(b1:b3)";
        int sheetNum = 0;
        int rowNum = 0;
        int cellNum = 0;
        writeFormulaCell(formula, sheetNum, rowNum, cellNum);
        assertModifiedFormula(sheetNum, rowNum, cellNum, formula.toUpperCase());
    }

    @Test
    public void writeFormulaOverExistingOne() throws Exception {
        String formula = "B4+D1+100";
        int sheetNum = 0;
        int rowNum = 5;
        int cellNum = 1;
        writeFormulaCell(formula, sheetNum, rowNum, cellNum);
        assertModifiedFormula(sheetNum, rowNum, cellNum, formula);
    }

    private void writeFormulaCell(String formula, int sheetNum, int rowNum, int cellNum) throws IOException {
        OutputStream resultStream = new FileOutputStream(MODIFIED_FILE_NAME);
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        if( row == null ){
            row = sheet.createRow(rowNum);
        }
        Cell cell = row.getCell(cellNum);
        if( cell == null ){
            cell = row.createCell(cellNum);
        }
        cell.setCellFormula( formula );
        workbook.write( resultStream );
        resultStream.close();
    }

    private void assertModifiedFormula(int sheetNum, int rowNum, int cellNum, String formula) throws Exception {
        InputStream inputStream = new FileInputStream(MODIFIED_FILE_NAME);
        Workbook reopenedWorkbook = WorkbookFactory.create( inputStream );
        Sheet sheet = reopenedWorkbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(cellNum);
        assertEquals("Written and read Formula values are different", formula, cell.getCellFormula());
        inputStream.close();
    }

    private void assertFormula(int sheetNum, int rowNum, int cellNum, String formula, String assertMessage){
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(cellNum);
        assertEquals(assertMessage, formula.toUpperCase(),
                cell.getCellFormula().toUpperCase());
    }

    @Test
    public void readFormulaValue(){
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(3);
        Cell cell = row.getCell(1);
        assertEquals("Formula result read incorrectly", 6, cell.getNumericCellValue(), 1e-6); 
    }
}
