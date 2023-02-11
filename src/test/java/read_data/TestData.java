package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class TestData {
    File excelFile = new File("src/test/resources/TestData.xlsx");
    FileInputStream fileInputStream;
    XSSFWorkbook workbook;
    XSSFSheet sheet;

    @Before
    public void setup() throws IOException {
        fileInputStream = new FileInputStream(excelFile);
        workbook = new XSSFWorkbook(fileInputStream);
        sheet = workbook.getSheet("Sheet1");
    }

    @Test
    public void getAllDataTest(){
        for (int i = sheet.getFirstRowNum();i <= sheet.getLastRowNum();i++){
            XSSFRow row = sheet.getRow(i);
            System.out.print("| ");
            for (int j = row.getFirstCellNum(); j< row.getLastCellNum();j++){
                System.out.print(row.getCell(j)+" | ");
            } System.out.println();
        }
    }
    @Test
    public void getBusinessTypeData(){
        String businessType = "BusinessType";
        XSSFRow row = sheet.getRow(0);
        int index = -1;

        for (int i = row.getFirstCellNum(); i < row.getLastCellNum();i++){
            XSSFCell cell = row.getCell(i);
                    if (cell.getStringCellValue().equalsIgnoreCase(businessType)){
                    index = i;
                }
            }
        for (int i = sheet.getFirstRowNum();i <= sheet.getLastRowNum();i++){
            XSSFRow row1 = sheet.getRow(i);
            System.out.println(row1.getCell(index));
        }
    }
}
