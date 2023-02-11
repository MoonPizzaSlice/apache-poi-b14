package readData;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class TestData {
    File file = new File("src/test/resources/TestData.xlsx");
    FileInputStream fileInputStream;
    XSSFWorkbook workbook;
    XSSFSheet sheetAt;

    @Before
    public void setup() throws IOException {
        fileInputStream = new FileInputStream(file);
        workbook = new XSSFWorkbook(fileInputStream);
        sheetAt = workbook.getSheetAt(0);
    }

    @Test
    public void allData() throws IOException {
        for (int i = sheetAt.getFirstRowNum(); i < sheetAt.getLastRowNum(); i++) {
            XSSFRow row = sheetAt.getRow(i);
            System.out.print("| ");
            for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
                XSSFCell cell = row.getCell(j);
                System.out.print(cell + " | ");
            }
            System.out.println();
        }
    }


    @Test
    public void columnData(){
//        for (int i = sheetAt.getFirstRowNum(); i<sheetAt.getLastRowNum(); i++){
//            XSSFRow row = sheetAt.getRow(i);
//            XSSFCell cell = row.getCell(row.getLastCellNum() - 3);
//            for (int k = row.getLastCellNum()-2; k< sheetAt.getLastRowNum(); k++){
//                System.out.println(cell);
//        }

        String columnName = "Construction";
        XSSFRow row = sheetAt.getRow(0);
        int index = -1;
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            XSSFCell tempCell = row.getCell(i);
            if (tempCell.getStringCellValue().equalsIgnoreCase(columnName)){
                index=i;
            }
        }

        if (index<0){
            throw  new RuntimeException();
        }

        for (int i = sheetAt.getFirstRowNum() ; i <=sheetAt.getLastRowNum(); i++) {
            XSSFRow tempRow = sheetAt.getRow(i);
            System.out.println(tempRow.getCell(index));
        }


    }
    }

