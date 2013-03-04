package org.poitest;

import org.apache.poi.ss.usermodel.*;

import static org.junit.Assert.assertEquals;
import static junit.framework.Assert.assertNull;
import static junit.framework.Assert.assertNotNull;

/**
 * Date: Apr 15, 2009
 *
 * @author Leonid Vysochyn
 */
public class Util {
    static void assertCellsEquality(Cell srcCell, Cell destCell){
        assertEquals("Cell types are not equal", srcCell.getCellType(), destCell.getCellType());
        assertEquals("Cell toString() are not equal", srcCell.toString(), destCell.toString());
        assertEquals("CellStyles are not equal", srcCell.getCellStyle(), destCell.getCellStyle());
        assertCellComments(srcCell.getCellComment(), destCell.getCellComment());
        assertEquals("Cell hyperlinks are not equal", srcCell.getHyperlink(), destCell.getHyperlink());
    }

    private static void assertCellComments(Comment comment1, Comment comment2) {
        if( comment1 == null ){
            assertNull( comment2 );
        }else{
            assertNotNull( comment2 );
            assertEquals("Comment strings are not equal", comment1.getString(), comment2.getString());
        }
    }

    public static void copyCells(Cell srcCell, Cell destCell) {
        XlsCellInfo cellInfo = new XlsCellInfo();
        cellInfo.readCell( srcCell );
        cellInfo.writeToCell( destCell );
    }

    public static void writeCellToRow(Row row, XlsCellInfo cellInfo) {
        int colIndex = cellInfo.getColIndex();
        Cell cell = row.getCell( colIndex );
        if( cell == null ){
            cell = row.createCell( colIndex );
        }
        cellInfo.writeToCell( cell );
    }

    public static XlsCellInfo readCell(Sheet sheet, int rowNum, int cellNum) {
        XlsCellInfo cellInfo = new XlsCellInfo();
        Row row = sheet.getRow( rowNum );
        Cell cell = row.getCell( cellNum );
        cellInfo.readCell( cell );
        return cellInfo;
    }

    public static void setCellComment(Cell cell, String commentText, String commentAuthor, ClientAnchor anchor){
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        Drawing drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        if( anchor == null ){
            anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex() + 1);
            anchor.setCol2(cell.getColumnIndex() + 3);
            anchor.setRow1(cell.getRowIndex());
            anchor.setRow2(cell.getRowIndex() + 2);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(commentAuthor != null ? commentAuthor : "");
        cell.setCellComment( comment );
    }
}
