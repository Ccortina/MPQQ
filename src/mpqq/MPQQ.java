/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mpqq;

import java.awt.Color;
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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

class MPQQProcTab6Helper{
    private int mpqqRowNumber;
    private XSSFWorkbook workbook;

    public MPQQProcTab6Helper(int mpqqRowNumber, XSSFWorkbook workbook) {
        this.mpqqRowNumber = mpqqRowNumber;
        this.workbook = workbook;
    }

    public int getMpqqRowNumber() {
        return mpqqRowNumber;
    }

    public void setMpqqRowNumber(int mpqqRowNumber) {
        this.mpqqRowNumber = mpqqRowNumber;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }
    
    
}

public class MPQQ {
    //Tab index for REFERENCE FILE
    //private static final int TRACKER = 0;
    private static final int USE_FOR_TAB1 = 1;
    private static final int USE_FOR_TAB6_1 = 2;
    private static final int USE_FOR_TAB6_2 = 3;
    
    //Columns number for reference file "Use for Tab 1"
    private static final int T2PEPSICO_STOCK_CODE = 3;
    private static final int T2PEPSICO_SUPPLIER_SITE_NAME = 7;

    //Columns number for reference file "Use for Tab 6(1)"

    
    //Tab index for MPQQ Template file
    private static final int MPQQ_TAB_SUPPLIER_BASIC_INFO = 1;
    private static final int MPQQ_TAB_TEST_DATA = 7;
    
    //Number of row per ingredient defined in MPQQ tab 6
    private static final int TEST_DATA_TAB_ROW_LIMIT = 15;
    private static final int MPQQ_TAB6_STARTING_COLUMN = 5;
    
    //Parameters that need concatenation of their description
    private static final String[] PARAMS_FOR_CONCAT = new String[]{"Particle Distribution",
                                            "Part. Dist.-Pan","Fatty Acid Composition"};
    
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
     
    /**
     * This method returns the a sheet grouped in a hasmap based on a column value
     * as Key.
     * @param sheet Receives the XSSFSheet for processing
     * @param keyCol    The column that will be used as key for the Map
     * @return returnMap a Map using keyCol and having a list of list to represent the sheet
     */
    private static Map<String,List< List<String> > > createMapFromSheet(XSSFSheet sheet,int keyCol){
        //The Map with a list of lists
        Map<String,List<List<String>>> returnMap = new HashMap<>();
        List<List<String>> ingredientRows;
        DataFormatter df = new DataFormatter();
        
        //Travel the sheet skipping the first row
        for(int rowNum=1; rowNum <= sheet.getLastRowNum(); rowNum++){
            Row row = sheet.getRow(rowNum);
            //Traverse the cell of the row
            Iterator<Cell> cellIterator = row.cellIterator();
            //Reset ingredient cell list
            List<String> ingredientCells = new ArrayList<>();
            while( cellIterator.hasNext() ){
                Cell cell = cellIterator.next();
                //Add to the List
                ingredientCells.add(df.formatCellValue(cell));
            }
            //If the ingredient already exist retrive the list
            if( returnMap.containsKey(ingredientCells.get(keyCol)) ){
                ingredientRows = returnMap.get( ingredientCells.get(keyCol) );
            }else{
                ingredientRows = new ArrayList<>();
            }
            //Add to the second list
            ingredientRows.add(ingredientCells);
            //Add to the Map
            returnMap.put(ingredientCells.get(keyCol), ingredientRows);    
        }  
        return returnMap;
    }
    
    /**
     * This method process the Tab "Use for tab 6 (2)" and returns a map for easy
     * search of classes , subclasses and their parameters.
     * @param sheet The sheet from where info will be extracted
     * @param ingClassCol   The Column in the sheet for the class(based index 0)
     * @param ingSubClassCol    The column number in the sheet for the subclass (based index 0)
     * @return return the map with the parameters
     */
    private static Map<String,Map<String,List<String> > > createMapFromSheet(XSSFSheet sheet,
                                                                            int ingClassCol,int ingSubClassCol){
        
        Map<String,Map<String,List<String> > > returnMap = new TreeMap<>(String.CASE_INSENSITIVE_ORDER); //Class with subclasses List
        
        DataFormatter df = new DataFormatter();
        
        Pattern regex = Pattern.compile("(\\*){2}$");
        
        //Iterate the sheet rows
        for(int rowNum=1; rowNum <= sheet.getLastRowNum(); rowNum++){
            Row row = sheet.getRow(rowNum);
            Map<String,List<String>> subclassAndParams = new TreeMap<>(String.CASE_INSENSITIVE_ORDER); //Subclass with its columns values
            String ingClass = df.formatCellValue(row.getCell(ingClassCol)); //Get the Ingredient Class
            String cleanClass;
            List<String> parameters = new ArrayList<>();
            Matcher matches = regex.matcher( ingClass );

            //Clean class text
            if(matches.find()){
                //The Class ends with ** , remove
                cleanClass = ingClass.
                        substring(0,ingClass.length()-2).
                        replaceAll("\\s+$", "");

            }else{
                cleanClass = ingClass.replaceAll("\\s+$", "");
            }
            
            if( returnMap.containsKey(cleanClass) ){
                //If class already exists
                subclassAndParams = returnMap.get(cleanClass);
                if( subclassAndParams.containsKey(df.formatCellValue(row.getCell(ingSubClassCol))) ){
                    //Subclass already present -- This should really happen
                    System.err.print("Subclass "+
                            df.formatCellValue(row.getCell(ingSubClassCol)) +
                            "seems to apper twice in the Class "+cleanClass);
                    System.exit(0);
                }else{
                    for( int cellIdx=ingSubClassCol+1; cellIdx <= row.getLastCellNum(); cellIdx++ ){
                        //Get the paramters
                        if(!df.formatCellValue(row.getCell(cellIdx)).isEmpty()){
                            String []paramsInCell = df.formatCellValue(row.getCell(cellIdx)).split("[\\n,]+");
                            parameters.addAll(Arrays.asList(paramsInCell));
                        }
                    }
                    //Add list to Subclass
                    subclassAndParams.put(df.formatCellValue(row.getCell(ingSubClassCol)), parameters);
                }
                
            }else{
                //If first class encounter -- Subclass cant exist yet
                for( int cellIdx=ingSubClassCol+1; cellIdx <= row.getLastCellNum(); cellIdx++ ){
                    //Get the paramters
                    if(!df.formatCellValue(row.getCell(cellIdx)).isEmpty()){
                        String []paramsInCell = df.formatCellValue(row.getCell(cellIdx)).split("[\\n,]+");
                        parameters.addAll(Arrays.asList(paramsInCell)); 
                    }
                }
                //Add list to Subclass
                subclassAndParams.put(df.formatCellValue(row.getCell(ingSubClassCol)), parameters);
            }
            returnMap.put(cleanClass, subclassAndParams);
        }
        return returnMap;
    }
    
    /*
     * This method process the information for Tab 1 
     * 1. Supplier Basic Info using Tracker tab from
     * reference file.
     * @param mpqqWB The workbook of the mpqq where the info will be written
     */
    private static XSSFWorkbook procTab1(Row referenceCurrentRow, 
                                            XSSFWorkbook mpqqWB,
                                            int mpqqCurrentRow){
        
        XSSFSheet tab1 = mpqqWB.getSheetAt(MPQQ_TAB_SUPPLIER_BASIC_INFO);
        DataFormatter df = new DataFormatter();

        int colIdx = 1;
        Iterator<Cell> cellIterator = referenceCurrentRow.cellIterator();
        while(cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            Cell currentCell = checkRowCellExists(tab1,mpqqCurrentRow,colIdx);
            switch(cell.getColumnIndex()){
                case 3: case 4:case 5: case 6: case 7: case 8: case 9:
                    currentCell.setCellValue(df.formatCellValue(cell));
                    //Go to next Column
                    colIdx++;
                    break;    
                default:
            }
        }
        return mpqqWB;
    }
    
    /*
     * This method fills the Tab 6 . Test Data on the mpqq document.
    */
    private static MPQQProcTab6Helper procTab6(String pepsicoStockCode,
                                            XSSFWorkbook mpqqWB,
                                            int mpqqStartingRow,
                                            Map<String,Map<String,List<String>>>map2,
                                            Map<String,List<List<String>>>map,
                                            int evenRow){
        
        XSSFSheet tab6 = mpqqWB.getSheetAt(MPQQ_TAB_TEST_DATA);

        List<List<String>> ingredientRows = map.get(pepsicoStockCode);
        List<String> ingredientParams = new ArrayList<>();
        
        //mpqqcurrentRow has the index for the current row meanwhile mpqqRowCounter 
        //has the numberof rows for that ingredient
        int mpqqCurrentRow = mpqqStartingRow,mpqqCurCol;

        XSSFColor myColor = new XSSFColor(Color.decode("#FFFF99"));
        XSSFCellStyle style = mpqqWB.createCellStyle();
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setFillForegroundColor(myColor);
        
        //Calculate if it will need additional params rows
        if( ingredientRows != null){
            if( ingredientRows.size() > 0){
                //Rows from Tab1
                int totalRows = ingredientRows.size();
                //Rows from tab 2
                String ingClass = ingredientRows.get(0).get(2);
                String ingSubClass = ingredientRows.get(0).get(3);
                
                Map<String,List<String>> ingClassMap = map2.get(ingClass);
                
                if(ingClassMap != null ){
                    if( ingClassMap.containsKey(ingSubClass)){
                        totalRows += map2.get(ingClass).get(ingSubClass).size();
                    }else{
                        System.err.println("The Sub-Class "+ingSubClass+" could not be found. Check reference document.");
                    }
                }else{
                    System.err.println("The Class "+ingClass+" for Ingredient "+ pepsicoStockCode +" could not be found. Check reference document.");
                }
                
                if(totalRows > TEST_DATA_TAB_ROW_LIMIT){
                    //Move the rows
                    tab6.shiftRows(mpqqCurrentRow+TEST_DATA_TAB_ROW_LIMIT, tab6.getLastRowNum()+1, totalRows-TEST_DATA_TAB_ROW_LIMIT);
                    //Copy first 4 columns of formulas
                    for(int i=1; i<=totalRows-TEST_DATA_TAB_ROW_LIMIT; i++){                       
                        for(int j=1; j<=4 ; j++){
                            Cell newCell = checkRowCellExists(tab6,mpqqCurrentRow+TEST_DATA_TAB_ROW_LIMIT+i-1,j);
                            newCell.setCellFormula(tab6.getRow(mpqqStartingRow).getCell(j).getCellFormula());
                            newCell.setCellStyle(style);
                        }
                    }
                }
                //Insert new values
                for(List<String> ingredientCells: ingredientRows){
                    mpqqCurCol = MPQQ_TAB6_STARTING_COLUMN;
                    
                    for(int cellIdx = 2; cellIdx <= 7; cellIdx++){
                        Cell mpqqCurCell = checkRowCellExists(tab6,mpqqCurrentRow,mpqqCurCol);
                        if( cellIdx == 7 && ingredientCells.size() >= 9 ){
                            String paramDesc =  ingredientCells.get(cellIdx)+"-"+ingredientCells.get(cellIdx+1);
                            if( !ingredientParams.contains(paramDesc) ){
                                //Set on cell and add to list
                                if(Arrays.asList(PARAMS_FOR_CONCAT).contains(ingredientCells.get(cellIdx))){
                                   //If the param must be concatenated with the description 
                                   mpqqCurCell.setCellValue(paramDesc); 
                                }else{
                                   mpqqCurCell.setCellValue(ingredientCells.get(cellIdx));
                                }
                                ingredientParams.add(paramDesc);
                            }
                        }else{
                            mpqqCurCell.setCellValue(ingredientCells.get(cellIdx));
                        }
                        mpqqCurCol++;
                    }
                    for(int i=1; i<= 18; i++){
                        Cell stylecell = checkRowCellExists(tab6, mpqqCurrentRow, i);
                        stylecell.setCellStyle(style);
                    }
                    mpqqCurrentRow++;
                }
            }
        }else{
            System.err.println("Ingredient "+ pepsicoStockCode + " was not found on reference document.");
        }
        
        return new MPQQProcTab6Helper(mpqqCurrentRow, mpqqWB);
    }
    
 
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        try{
            //Read Reference File
            FileInputStream referenceFile = new FileInputStream(
                    new File("C:\\Users\\admin\\Documents\\MPQQ\\iRef.xlsx"));
            XSSFWorkbook reference = new XSSFWorkbook(referenceFile);
            reference.close();
            
            //Load MPQQ template
            FileInputStream mpqqFile = new FileInputStream(
                    new File("C:\\Users\\admin\\Documents\\MPQQ\\MPQQ_Template.xlsm"));
            XSSFWorkbook mpqq = new XSSFWorkbook( mpqqFile );
            mpqqFile.close();
            //First row to consider from Reference File based on 0 index
            int referenceStartRow = 156;
            
            XSSFSheet trackerTab = reference.getSheetAt(USE_FOR_TAB1);
            
            //MPQQ first row  
            int mpqqTab1CurrentRow = 11,mpqqTab1FirstRow = 11;
            int mpqqTab6CurrentRow = 9,mpqqTab6FirstRow = 9;  
            String currentSupplier ="";
            DataFormatter df = new DataFormatter();
            
            //Get all the valid rows in the range, not empty nor hidden rows.
            List<Row> validRows = new ArrayList<>();
            for(int refCurRow = referenceStartRow; refCurRow <= trackerTab.getLastRowNum();refCurRow++){
                Row row = trackerTab.getRow(refCurRow); //Current row in the Reference file
                
                //Check if row is visible
                if( !row.getZeroHeight() || (row.isFormatted() && row.getRowStyle().getHidden())){
                    //Valid - add it to the list
                    validRows.add(row);
                }
            }
            
            //Process the valid rows
            if(validRows.size() > 0 ){
                for(int i =0; i < validRows.size(); i++ ){
                   Row curRefRow = validRows.get(i);
                   currentSupplier = df.formatCellValue(curRefRow.getCell(T2PEPSICO_SUPPLIER_SITE_NAME));
                   
                    mpqq = procTab1(curRefRow, mpqq, mpqqTab1CurrentRow);
                    
                    Map<String,List<List<String>>> map = createMapFromSheet(reference.getSheetAt(USE_FOR_TAB6_1),1);
                    Map<String,Map<String,List<String>>> tab62Map = createMapFromSheet(reference.getSheetAt(USE_FOR_TAB6_2),0,1);

                    MPQQProcTab6Helper auxiliarClass = procTab6(df.formatCellValue(curRefRow.getCell(T2PEPSICO_STOCK_CODE)),
                                                        mpqq,
                                                        mpqqTab6CurrentRow,
                                                        tab62Map,
                                                        map,
                                                        i&2);
                    mpqq = auxiliarClass.getWorkbook();
                    
                    //Jump Next Row on the MQPP
                    mpqqTab1CurrentRow++;
                    mpqqTab6CurrentRow = auxiliarClass.getMpqqRowNumber();
                    
                    //Check if next supplier changes or is this last row?
                    if( i+1 == validRows.size() ){
                        XSSFFormulaEvaluator.evaluateAllFormulaCells(mpqq);
                        try{
                            FileOutputStream outFile = new FileOutputStream(
                                    new File("C:\\Users\\admin\\Documents\\MPQQ\\Output\\"+
                                            currentSupplier.replaceAll("[^a-zA-Z]+", "")+".xlsm"));
                            mpqq.write(outFile);
                            outFile.close();
                        }catch(Exception w){
                            w.printStackTrace(System.out);
                        }
                    }else{
                        //Check if the supplier changes
                        String nextSupplier = df.formatCellValue(validRows.get(i+1).getCell(T2PEPSICO_SUPPLIER_SITE_NAME));
                        if( !currentSupplier.equalsIgnoreCase(nextSupplier) ){
                            XSSFFormulaEvaluator.evaluateAllFormulaCells(mpqq);
                            
                            try{
                                FileOutputStream outFile = new FileOutputStream(
                                        new File("C:\\Users\\admin\\Documents\\MPQQ\\Output\\"+
                                                currentSupplier.replaceAll("[^a-zA-Z]+", "")+".xlsm"));
                                mpqq.write(outFile);
                                outFile.close();
                            }catch(Exception w){
                                w.printStackTrace(System.out);
                            }
                            //Reload Template
                            mpqqFile = new FileInputStream(
                                    new File("C:\\Users\\admin\\Documents\\MPQQ\\MPQQ_Template.xlsm"));
                            mpqq = new XSSFWorkbook( mpqqFile );
                            mpqqFile.close();
                            mpqqTab1CurrentRow = mpqqTab1FirstRow;
                            mpqqTab6CurrentRow = mpqqTab6FirstRow;
                        }
                    }

                }
            }
            
        } catch ( FileNotFoundException e){
            e.printStackTrace();
        } catch ( IOException e ){
            e.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }
        
        System.out.println("Testing");
    }
    
}
