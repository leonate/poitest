package org.poitest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Date;

/**
 * Date: Apr 15, 2009
 *
 * @author Leonid Vysochyn
 */
public class XlsCellInfo {

    RichTextString richTextString;
    boolean booleanValue;

    int cellType;
    private Date dateValue;
    private double doubleValue;
    private String formula;
    private CellStyle style;
    private Comment comment;
    private Hyperlink hyperlink;
    private byte errorValue;
    private int rowIndex;
    private int colIndex;
    private RichTextString commentString;
    private String commentAuthor;

    public int getColIndex() {
        return colIndex;
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        cellType = cell.getCellType();
        comment = cell.getCellComment();
        if( comment != null ){
            commentString = cell.getCellComment().getString();
            commentAuthor = cell.getCellComment().getAuthor();
        }
        hyperlink = cell.getHyperlink();
        colIndex = cell.getColumnIndex();
        rowIndex = cell.getRowIndex();
    }

    private void readCellContents(Cell cell) {
        switch( cell.getCellType() ){
            case Cell.CELL_TYPE_STRING:
                richTextString = cell.getRichStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                  dateValue = cell.getDateCellValue();
                } else {
                  doubleValue = cell.getNumericCellValue();
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                formula = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_ERROR:
                errorValue = cell.getErrorCellValue();
                break;
        }
    }

    private void readCellStyle(Cell cell) {
        style = cell.getCellStyle();
    }

    public void writeToCell(Cell cell){
        updateCellGeneralInfo(cell);
        updateCellContents( cell );
        updateCellStyle( cell );
    }

    private void updateCellGeneralInfo(Cell cell) {
        cell.setCellType( cellType );
        updateCellComment( cell );
        if( hyperlink != null ){
            cell.setHyperlink( hyperlink );
        }
    }

    private void updateCellComment(Cell cell) {
        if( comment != null ){
            Workbook wb = cell.getSheet().getWorkbook();
            Sheet sheet = cell.getSheet();
            Drawing drawing = sheet.createDrawingPatriarch();
            CreationHelper factory = wb.getCreationHelper();
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(4);
            anchor.setCol2(6);
            anchor.setRow1(2);
            anchor.setRow2(5);
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(commentString);
            comment.setAuthor(commentAuthor);
            cell.setCellComment( comment );
        }
    }


    private void updateCellContents(Cell cell) {
        switch( cellType ){
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue( richTextString );
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue( booleanValue );
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if( dateValue != null ){
                    cell.setCellValue( dateValue );
                }else{
                    cell.setCellValue( doubleValue );
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellFormula(formula);
                break;
            case Cell.CELL_TYPE_ERROR:
                cell.setCellErrorValue( errorValue );
                break;
        }
    }

    private void updateCellStyle(Cell cell) {
        cell.setCellStyle( style );
    }

}
