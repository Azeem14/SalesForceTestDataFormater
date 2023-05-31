//package org.example;
//
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.util.*;
//
//public class TestDataExtracter {
//
//    public static void main(String args[]){
//        try {
//            LinkedHashMap<String,String> testData= new LinkedHashMap<>();
//            String columnName="";
//            String value="";
//            boolean br=false;
//            boolean flag=false;
//            boolean teststeps=false;
//            LinkedHashMap<String, List<String>> testSteps = new LinkedHashMap<>();
//            FileInputStream file = new FileInputStream(
//                    new File("src/main/resources/Testdata/Script_from_QualiGen_Sample.xlsx"));
//            XSSFWorkbook workbook = new XSSFWorkbook(file);
//            XSSFSheet sheet = workbook.getSheetAt(0);
//            Iterator<Row> rowIterator = sheet.iterator();
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                Iterator<Cell> cellIterator
//                        = row.cellIterator();
//                 int index=0;
//                 String step="";
//                 String description="";
//                 String expectedResults="";
//                 while (cellIterator.hasNext() && br) {
//
//                     Cell cell = cellIterator.next();
//                     if(cell.getCellType().name().equals("NUMERIC")){
//                         teststeps=true;
//                        step="Step "+(int)cell.getNumericCellValue();
//                     }
//                     if(cell.getCellType().name().equals("STRING")&&teststeps){
//                         index++;
//                         if(index==1){ description=cell.getStringCellValue().trim();}
//                         if(index==2){
//                             expectedResults=cell.getStringCellValue().trim();
//                         }
//
//                     }
//                 }
//                 if(teststeps){
//                 testSteps.put(step,new ArrayList<String>(Arrays.asList(description,expectedResults)));}
//
//                while (cellIterator.hasNext() && br==false) {
//
//                    Cell cell = cellIterator.next();
//
//                    if(cell.getCellType().name().equals("STRING")){
//                        System.out.println(cell.getStringCellValue());
//
//                        if(cell.getStringCellValue().split(":").length==2&&flag==false) {
//                            testData.put(cell.getStringCellValue().split(":")[0].trim(), cell.getStringCellValue().split(":")[1].trim());
//                        }
//                        if(cell.getStringCellValue().split(":").length==1&&flag==false){
//                           // testData.put(cell.getStringCellValue().split(":")[0].trim(),"");
//                            columnName=cell.getStringCellValue().split(":")[0].trim();
//                            if(columnName.equals("Test Steps")){
//                                br=true;
//                                break;
//                            }
//                            flag=true;
//                            continue;
//                        }
//                        if(flag){
//                                if(!cell.getStringCellValue().contains("Test Steps")) {
//                                    if(value.equals("")) {
//                                        value = cell.getStringCellValue();
//                                    }
//                                    else {
//                                        value = value + "\r\n"+" " + cell.getStringCellValue();
//                                    }
//
//                                }
//                        }
//                        if(cell.getStringCellValue().contains(":")&&cell.getStringCellValue().split(":").length==1&&flag) {
//                            testData.put(columnName,value);
//                            columnName="";
//                            value="";
//                            columnName= cell.getStringCellValue().split(":")[0].trim();
//                            flag=false;
//                            if(cell.getStringCellValue().contains("Test Steps")){
//                                br=true;
//                                testData.put("Test Step Number","");
//                                testData.put("Test Step description","");
//                                testData.put("Expected Results","");
//                                break;
//                            }
//                        }
//                    }
//                }
//
//            }
//            System.out.println("---------------------------------------");
//            System.out.println(testData);
//            System.out.println(testSteps);
//            file.close();
//            XSSFWorkbook resultworkbook = new XSSFWorkbook();
//            XSSFSheet resultsheet = resultworkbook.createSheet(testData.get("Test Script ID"));
//            int rownum = 0;
//            Row row = resultsheet.createRow(rownum++);
//            int cellnum = 0;
//            Cell cell;
//            for(Map.Entry<String, String> entry:testData.entrySet()){
//                cell = row.createCell(cellnum++);
//                cell.setCellValue(entry.getKey());
//
//            }
//            row = resultsheet.createRow(rownum++);
//            cellnum = 0;
//            for(Map.Entry<String, String> entry:testData.entrySet()){
//                cell = row.createCell(cellnum++);
//                cell.setCellValue(entry.getValue());
//            }
//            for(Map.Entry<String, List<String>> entry:testSteps.entrySet()){
//                cell = row.createCell(5);
//                cell.setCellValue(entry.getKey());
//                cell = row.createCell(6);
//                cell.setCellValue(entry.getValue().get(0));
//                cell = row.createCell(7);
//                cell.setCellValue(entry.getValue().get(1));
//                row = resultsheet.createRow(rownum++);
//            }
//
//            try {
//                FileOutputStream out = new FileOutputStream(
//                        new File("src/main/resources/EndData/EndtestDataResult.xlsx"));
//                resultworkbook.write(out);
//                out.close();
//                System.out.println(
//                        "EndtestDataResult.xlsx written successfully on disk.");
//            }
//            catch (Exception e) {
//                e.printStackTrace();
//            }
//
//
//        }
//        catch (Exception e) {
//            e.printStackTrace();
//        }
//
//    }
//
//}/// single Version


// adding multiple testcases to
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
        List<LinkedHashMap> mas=new ArrayList<>();
        LinkedHashMap<LinkedHashMap<String, String>, LinkedHashMap<String, List<String>>> master = new LinkedHashMap<>();
        LinkedHashMap<String, String> testData;
        LinkedHashMap<String, List<String>> testSteps;
        File directoryPath = new File("src/main/resources/Testdata");
        try {
            for(String filename:directoryPath.list()){
                testData = new LinkedHashMap<>();
                testSteps = new LinkedHashMap<>();
            String columnName = "";
            String value = "";
            boolean br = false;
            boolean flag = false;
            boolean teststeps = false;
            FileInputStream file = new FileInputStream(
                    new File("src/main/resources/Testdata/"+filename));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator
                        = row.cellIterator();
                int index = 0;
                String step = "";
                String description = "";
                String expectedResults = "";
                while (cellIterator.hasNext() && br) {

                    Cell cell = cellIterator.next();
                    if (cell.getCellType().name().equals("NUMERIC")) {
                        teststeps = true;
                        step = "Step " + (int) cell.getNumericCellValue();
                    }
                    if (cell.getCellType().name().equals("STRING") && teststeps) {
                        index++;
                        if (index == 1) {
                            description = cell.getStringCellValue().trim();
                        }
                        if (index == 2) {
                            expectedResults = cell.getStringCellValue().trim();
                        }

                    }
                }
                if (teststeps) {
                    testSteps.put(step, new ArrayList<String>(Arrays.asList(description, expectedResults)));
                }

                while (cellIterator.hasNext() && br == false) {

                    Cell cell = cellIterator.next();

                    if (cell.getCellType().name().equals("STRING")) {
                        System.out.println(cell.getStringCellValue());

                        if (cell.getStringCellValue().split(":").length == 2 && flag == false) {
                            testData.put(cell.getStringCellValue().split(":")[0].trim(), cell.getStringCellValue().split(":")[1].trim());
                        }
                        if (cell.getStringCellValue().split(":").length == 1 && flag == false) {
                            // testData.put(cell.getStringCellValue().split(":")[0].trim(),"");
                            columnName = cell.getStringCellValue().split(":")[0].trim();
                            if (columnName.equals("Test Steps")) {
                                br = true;
                                break;
                            }
                            flag = true;
                            continue;
                        }
                        if (flag) {
                            if (!cell.getStringCellValue().contains("Test Steps")) {
                                if (value.equals("")) {
                                    value = cell.getStringCellValue();
                                } else {
                                    value = value + "\r\n" + " " + cell.getStringCellValue();
                                }

                            }
                        }
                        if (cell.getStringCellValue().contains(":") && cell.getStringCellValue().split(":").length == 1 && flag) {
                            testData.put(columnName, value);
                            columnName = "";
                            value = "";
                            columnName = cell.getStringCellValue().split(":")[0].trim();
                            flag = false;
                            if (cell.getStringCellValue().contains("Test Steps")) {
                                br = true;
                                testData.put("Test Step Number", "");
                                testData.put("Test Step description", "");
                                testData.put("Expected Results", "");
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
            master.put(testData, testSteps);
            mas.add(testData);
            mas.add(testSteps);
            System.out.println("Interator");
        }
            System.out.println(master);



            XSSFWorkbook resultworkbook = new XSSFWorkbook();
            XSSFSheet resultsheet = resultworkbook.createSheet("TestData");
            int rownum = 0;
            Row row = resultsheet.createRow(rownum++);
            int cellnum = 0;
            Cell cell;
            List<String> columnHeaders=new ArrayList<>(Arrays.asList("Subject","Module","Sub Module","Test Script ID","Test Name","Test script Description","Prerequisites","Role","Test Data","Test Step Number","Test Step description","Expected Results","Business Priority","Requirement","Type"));

            for(String header: columnHeaders) {
                cell = row.createCell(cellnum++);
                cell.setCellValue(header);
            }
            int index=0;
            for(int i=0;i<mas.size();i++){
                if(i>=mas.size()/2){break;}
            //for(Map.Entry<LinkedHashMap<String,String>,LinkedHashMap<String,List<String>>> masterentry:master.entrySet()) {
//                testData = masterentry.getKey();
//                testSteps = masterentry.getValue();
                testData=mas.get(index++);
                testSteps=mas.get(index++);
                row = resultsheet.createRow(rownum++);
                for (Map.Entry<String, String> entry : testData.entrySet()) {
                    String columnhearder = entry.getKey();
                    String data = entry.getValue();
                    columnhearder = columnhearder.equals("Objective") ? "Test script Description" : columnhearder;
                    for (int k = 0; k < columnHeaders.size(); k++) {
                        Row row1 = resultsheet.getRow(0);
                        cell = row1.getCell(k);
                        if (cell.getStringCellValue().equals(columnhearder)) {
                            cell = row.createCell(k);
                            cell.setCellValue(data);
                        }
                    }
                }

                for (Map.Entry<String, List<String>> entry : testSteps.entrySet()) {
                    for (int j = 0; j < columnHeaders.size(); j++) {
                        Row row1 = resultsheet.getRow(0);
                        cell = row1.getCell(j);
                        if (cell.getStringCellValue().equals("Test Step Number")) {
                            Cell cell1 = row.createCell(j);
                            cell1.setCellValue(entry.getKey());
                        }
                        if (cell.getStringCellValue().equals("Test Step description")) {
                            Cell cell1 = row.createCell(j);
                            cell1.setCellValue(entry.getValue().get(0));
                        }
                        if (cell.getStringCellValue().equals("Expected Results")) {
                            Cell cell1 = row.createCell(j);
                            cell1.setCellValue(entry.getValue().get(1));
                        }

                    }
                    row = resultsheet.createRow(rownum++);
                }
            }

                try {
                    FileOutputStream out = new FileOutputStream(
                            new File("src/main/resources/EndData/EndtestDataResult.xlsx"));
                    resultworkbook.write(out);
                    out.close();
                    System.out.println(
                            "EndtestDataResult.xlsx written successfully on disk.");
                } catch (Exception e) {
                    e.printStackTrace();
                }

        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }

}
