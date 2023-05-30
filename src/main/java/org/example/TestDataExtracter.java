package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class TestDataExtracter {

    public static void main(String args[]){
        try {
            LinkedHashMap<String,String> testData= new LinkedHashMap<>();
            String columnName="";
            String value="";
            boolean br=false;
            boolean flag=false;
            boolean teststeps=false;
            LinkedHashMap<String, List<String>> testSteps = new LinkedHashMap<>();
            FileInputStream file = new FileInputStream(
                    new File("src/main/resources/Testdata/Script_from_QualiGen_Sample.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator
                        = row.cellIterator();
                 int index=0;
                 String step="";
                 String description="";
                 String expectedResults="";
                 while (cellIterator.hasNext() && br) {

                     Cell cell = cellIterator.next();
                     if(cell.getCellType().name().equals("NUMERIC")){
                         teststeps=true;
                        step="Step "+(int)cell.getNumericCellValue();
                     }
                     if(cell.getCellType().name().equals("STRING")&&teststeps){
                         index++;
                         if(index==1){ description=cell.getStringCellValue().trim();}
                         if(index==2){
                             expectedResults=cell.getStringCellValue().trim();
                         }

                     }
                 }
                 if(teststeps){
                 testSteps.put(step,new ArrayList<String>(Arrays.asList(description,expectedResults)));}

                while (cellIterator.hasNext() && br==false) {

                    Cell cell = cellIterator.next();

                    if(cell.getCellType().name().equals("NUMERIC")){
                        System.out.println(cell.getNumericCellValue());
                    }
                    if(cell.getCellType().name().equals("STRING")){
                        System.out.println(cell.getStringCellValue());

                        if(cell.getStringCellValue().split(":").length==2&&flag==false) {
                            testData.put(cell.getStringCellValue().split(":")[0].trim(), cell.getStringCellValue().split(":")[1].trim());
                        }
                        if(cell.getStringCellValue().split(":").length==1&&flag==false){
                           // testData.put(cell.getStringCellValue().split(":")[0].trim(),"");
                            columnName=cell.getStringCellValue().split(":")[0].trim();
                            if(columnName.equals("Test Steps")){
                                br=true;
                                break;
                            }
                            flag=true;
                            continue;
                        }
                        if(flag){
                                if(!cell.getStringCellValue().contains("Test Steps")) {
                                    if(value.equals("")) {
                                        value = cell.getStringCellValue();
                                    }
                                    else {
                                        value = value + "\r\n"+" " + cell.getStringCellValue();
                                    }

                                }
                        }
                        if(cell.getStringCellValue().contains(":")&&cell.getStringCellValue().split(":").length==1&&flag) {
                            testData.put(columnName,value);
                            columnName="";
                            value="";
                            columnName= cell.getStringCellValue().split(":")[0].trim();
                            flag=false;
                            if(cell.getStringCellValue().contains("Test Steps")){
                                br=true;
                                testData.put("Test Step Number","");
                                testData.put("Test Step description","");
                                testData.put("Expected Results","");
                                break;
                            }
                        }
                    }
                }

            }
            System.out.println("---------------------------------------");
            System.out.println(testData);
            System.out.println(testSteps);
            file.close();
            XSSFWorkbook resultworkbook = new XSSFWorkbook();
            XSSFSheet resultsheet = resultworkbook.createSheet(testData.get("Test Script ID"));
            int rownum = 0;
            Row row = resultsheet.createRow(rownum++);
            int cellnum = 0;
            Cell cell;
            for(Map.Entry<String, String> entry:testData.entrySet()){
                cell = row.createCell(cellnum++);
                cell.setCellValue(entry.getKey());

            }
            row = resultsheet.createRow(rownum++);
            cellnum = 0;
            for(Map.Entry<String, String> entry:testData.entrySet()){
                cell = row.createCell(cellnum++);
                cell.setCellValue(entry.getValue());
            }
            for(Map.Entry<String, List<String>> entry:testSteps.entrySet()){
                cell = row.createCell(5);
                cell.setCellValue(entry.getKey());
                cell = row.createCell(6);
                cell.setCellValue(entry.getValue().get(0));
                cell = row.createCell(7);
                cell.setCellValue(entry.getValue().get(1));
                row = resultsheet.createRow(rownum++);
            }

            try {
                FileOutputStream out = new FileOutputStream(
                        new File("src/main/resources/EndData/testDataResult"+testData.get("Test Script ID")+".xlsx"));
                resultworkbook.write(out);
                out.close();
                System.out.println(
                        "testDataResult.xlsx written successfully on disk.");
            }
            catch (Exception e) {
                e.printStackTrace();
            }


        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }

}
