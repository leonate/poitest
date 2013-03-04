package org.poitest;

import org.apache.poi.ss.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.*;

import static junit.framework.Assert.assertEquals;
import static junit.framework.Assert.assertNotNull;

/**
 * @author Leonid Vysochyn
 */
public class CommentTest {
    public static final String TEMPLATE_FILE_NAME = "storecopy.xls";
    public static final String MODIFIED_FILE_NAME = "target/comment.xls";

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
    public void createComment() throws IOException {
        String comment = "My comment";
        String author = "leo";
        int sheetNum = 0;
        int rowNum = 3;
        int cellNum = 1;
        writeCommentToCell(comment, author, sheetNum, rowNum, cellNum);
        assertComment(sheetNum, rowNum, cellNum, comment, author);
    }

    private void writeCommentToCell(String comment, String author, int sheetNum, int rowNum, int cellNum) throws IOException {
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
        Util.setCellComment(cell, comment, author, null);
        cell.setCellValue("Test cell");
        workbook.write( resultStream );
        resultStream.close();
    }

    private void assertComment(int sheetNum, int rowNum, int cellNum, String comment, String author){
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        assert row != null;
        Cell cell = row.getCell(cellNum);
        assert cell != null;
        Comment commentObj = cell.getCellComment();
        assert commentObj != null;
        assertEquals("Comment text was written incorrectly", comment, commentObj.getString().getString() );
        assertEquals("Comment author was written incorrectly", author, commentObj.getAuthor());
    }

}
