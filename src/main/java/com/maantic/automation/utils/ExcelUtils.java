package com.maantic.automation.utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    private ExcelUtils(){

    }

    public static List<Map<String,String>> getExcelData(String sheetName){
        List<Map<String,String>> list = null;

        FileInputStream fs;

        try{
            // System.out.println("Data File"+Constants.TEST_DATA_SHEET_PATH);
            fs= new FileInputStream(Constants.TEST_DATA_SHEET_PATH);
            XSSFWorkbook wb= new XSSFWorkbook(fs);
            XSSFSheet wSheet = wb.getSheet(sheetName);

            int lastRowNum= wSheet.getLastRowNum();
            int lastColNum= wSheet.getRow(0).getLastCellNum();

            Map<String,String> dataMap=null;
            list = new ArrayList<>();

            for(int i=1;i<=lastRowNum;i++){
                dataMap = new HashMap<>();
                for(int k=0;k<lastColNum;k++){
                    String key = wSheet.getRow(0).getCell(k).getStringCellValue();
                    String value = wSheet.getRow(i).getCell(k).getStringCellValue();
                    dataMap.put(key,value);
                }
                list.add(dataMap);
            }
            return list;
        }
        catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
