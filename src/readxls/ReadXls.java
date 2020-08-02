/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readxls;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author 62813
 */
public class ReadXls {

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String[] args) throws IOException {
        // TODO code application logic here
        
        readFromExcel("D:data.xls");
    }
    
    public static void readFromExcel(String urlexcel) throws FileNotFoundException, IOException{
        HSSFWorkbook myexcel = new HSSFWorkbook(new FileInputStream(urlexcel));
        HSSFSheet myexcelSheet = myexcel.getSheet("rubik");
        FormulaEvaluator formulaEv = myexcel.getCreationHelper().createFormulaEvaluator();
        
        for (Row row: myexcelSheet) {
            for(org.apache.poi.ss.usermodel.Cell cell : row){
                switch(formulaEv.evaluateInCell(cell).getCellType()){
                    
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        
                        break;
                        
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        
                        break;
                }
            }
            
            System.out.println("");
            myexcel.close();
            
        }
    }  
}
