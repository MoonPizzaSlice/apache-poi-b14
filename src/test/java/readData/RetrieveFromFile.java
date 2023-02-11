package readData;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class RetrieveFromFile {

    @Test
    public  void readFileTest() throws IOException {

        File excelFile = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream =new FileInputStream(excelFile); //reads content of the file and saves it in a storage( fileInputStream)

        XSSFWorkbook workbook =new XSSFWorkbook(fileInputStream);
        XSSFSheet page1= workbook.getSheet("Sheet1");
        XSSFRow row1 = page1.getRow(0);
        XSSFCell cell1 = row1.getCell(0);
        XSSFRow row2 = page1.getRow(1);
        XSSFCell cell2 = row2.getCell(0);
        System.out.println(cell1);
        System.out.println(cell2);
    }

    @Test
    public void getRowValuesTest() throws IOException {
        File excelFile= new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream =new FileInputStream(excelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        XSSFRow row0 = sheet1.getRow(0);
        for (int i = row0.getFirstCellNum(); i<row0.getLastCellNum(); i++){
            System.out.print(row0.getCell(i)+" | ");
        }
    }

    @Test
    public void getAllValues() throws IOException {
        File excelFile= new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream =new FileInputStream(excelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        for (int i = sheet1.getFirstRowNum(); i<sheet1.getLastRowNum(); i++){
            XSSFRow row = sheet1.getRow(i);
            System.out.print("| ");
            for (int j = row.getFirstCellNum(); j<row.getLastCellNum(); j++){
                XSSFCell cell = row.getCell(j);
                System.out.print(cell+ " | ");
            }
            System.out.println();
        }
    }


}
