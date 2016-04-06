/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mpqq;

import java.io.File;
/**
 *
 * @author 09168336
 */
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

public class MPQQ {
    //Tab index for REFERENCE FILE
    private static final int TRACKER = 0;
    private static final int USE_FOR_TAB1 = 1;
    private static final int USE_FOR_TAB6_1 = 2;
    private static final int USE_FOR_TAB6_2 = 3;
    
    //Columns number for reference file "Use for Tab 1"
    private static final int T2PEPSICO_STOCK_CODE = 3;
    private static final int T2PEPSICO_INGREDIENT_NAME = 4;
    private static final int T2PEPSICO_SUPPLIER_SITE_MATERIAL = 5;
    private static final int T2PEPSICO_SUPPLIER_SITE_CODE = 6;
    private static final int T2PEPSICO_SUPPLIER_SITE_NAME = 7;
    private static final int T2PEPSICO_SUPPLIER_MATERIAL_NAME = 8;
    private static final int T2PEPSICO_SUPPLIER_MATERIAL_CODE = 9;
    
    /**
     * This method will check if the index for the row
     * and column are valid and return the cell, if not
     * it will create row and cell based on index.
     */
    private static Cell checkRowCellExists(XSSFSheet currentSheet,int rowIndex, int colIndex){
        Row currentRow = currentSheet.getRow(rowIndex);
        if( currentRow == null){
            currentRow = currentSheet.createRow(rowIndex);
        }
        //Check if cell exists
        Cell currentCell = currentRow.getCell(colIndex);
        if( currentCell == null){
            currentCell = currentRow.createCell(colIndex);
        }
        return currentCell;
    }
    
    /*
     * This method process the information for Tab 1 
     * 1. Supplier Basic Info using Tracker tab from
     * reference file.
     * @param referenceWB The workbook from were the information is going to be copied
     * @param mpqqWB The workbook of the mpqq where the info will be written
     * @param firstRow The first row from where the program must start copying
     */
    private static XSSFWorkbook procTab1(XSSFWorkbook referenceWB, 
                                            XSSFWorkbook mpqqWB,
                                            int referenceFirstRow){
        
        XSSFSheet trackerTab = referenceWB.getSheetAt(USE_FOR_TAB1);
        XSSFSheet tab1 = mpqqWB.getSheetAt(1);
        //Iterator<Row> rowIterator = trackerTab.iterator();
        DataFormatter df = new DataFormatter();

        //MPQQ first row    
        int rowIdx = 11;
        
        for(int refCurRow = referenceFirstRow; refCurRow <= trackerTab.getLastRowNum();refCurRow++){
            Row row = trackerTab.getRow(refCurRow);
            
            //Check if row is visible
            if( !row.getZeroHeight() || (row.isFormatted() && row.getRowStyle().getHidden())){

                int colIdx = 1;
                //Iterate trough the Columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch(cell.getColumnIndex()){
                        case 3:
                            Cell currentCell = checkRowCellExists(tab1,rowIdx,colIdx);
                            currentCell.setCellValue(df.formatCellValue(row.getCell(T2PEPSICO_STOCK_CODE)));
                            
                            //Go to next Column
                            colIdx++;
                            break;
                        case 4:
                            
                            break;
                        default:
                    }
                }
                //Jump Next Row
                rowIdx++;
            }
        }
        return mpqqWB;
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        
        try{
            //Read Reference File
            FileInputStream referenceFile = new FileInputStream(
                    new File("C:\\Personal\\09168336\\Documents\\iRef.xlsx"));
            XSSFWorkbook reference = new XSSFWorkbook(referenceFile);
            reference.close();
            
            FileInputStream mpqqFile = new FileInputStream(
                    new File("C:\\Personal\\09168336\\Documents\\mpqq.xlsm"));
            XSSFWorkbook mpqq = new XSSFWorkbook( mpqqFile );
            mpqqFile.close();
            int referenceStartRow = 156;
            
            //Create a MPQQ for each supplier in the reference document
            String currentSupplier = reference.getSheetAt(USE_FOR_TAB1)
                    .getRow(referenceStartRow)
                    .getCell(T2PEPSICO_SUPPLIER_SITE_NAME)
                    .getStringCellValue();
            
            System.out.println(currentSupplier);
            
            mpqq = procTab1(reference, mpqq,referenceStartRow);
            FileOutputStream outFile = new FileOutputStream(new File("C:\\Personal\\09168336\\Documents\\mpqq.xlsm"));
            mpqq.write(outFile);
            outFile.close();
            
        } catch ( FileNotFoundException e){
            e.printStackTrace();
        } catch ( IOException e ){
            e.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }
        
        System.out.println("Testing");
    }
    
    private static void printsheet(XSSFSheet sheet){
        //Iterate through each rows from first sheet
    Iterator<Row> rowIterator = sheet.iterator();
    while(rowIterator.hasNext()) {
        Row row = rowIterator.next();
         
        //For each row, iterate through each columns
        Iterator<Cell> cellIterator = row.cellIterator();
        while(cellIterator.hasNext()) {
             
            Cell cell = cellIterator.next();
             
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    System.out.print(cell.getBooleanCellValue() + "\t\t");
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    System.out.print(cell.getNumericCellValue() + "\t\t");
                    break;
                case Cell.CELL_TYPE_STRING:
                    System.out.print(cell.getStringCellValue() + "\t\t");
                    break;
            }
        }
        System.out.println("");
    }
    }
    
}
