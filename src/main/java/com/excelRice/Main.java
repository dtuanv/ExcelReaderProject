package com.excelRice;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class Main {
    public static void main(String[] args) throws IOException {

        String excelFilePath = "T:\\Project\\Backend\\File\\nhanh.xlsx";
        String exportFile = "T:\\Project\\Backend\\File\\ExportData3.xlsx";

        ExcelFile excelFile = new ExcelFile();

//        excelFile.extracted(excelFilePath, exportFile);

        excelFile.getHeader(excelFilePath).forEach(header -> {
            System.out.println("header.name: "+header.name);
            System.out.println("index: "+header.index);
        });


        // Specify the Excel file path
        String excelToJson = "T:\\Project\\Backend\\File\\nhanh.xlsx";  // Path to your Excel file

        // Convert Excel to JSON
        List<Map<String, String>> excelData = readExcel(excelToJson);

        // Convert to JSON using Jackson
        ObjectMapper mapper = new ObjectMapper();
        String json = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(excelData);

        // Print JSON output
        System.out.println(json);
    }

    public static List<Map<String, String>> readExcel(String filePath) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();

        // Open the Excel file
        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);  // Reading the first sheet

        Iterator<Row> rowIterator = sheet.iterator();
        Row headerRow = rowIterator.next();  // First row as header

        // Get header values
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }

        // Read rows after the header
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Map<String, String> rowData = new HashMap<>();

            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.getCell(i);
                if (cell != null) {
                    rowData.put(headers.get(i), getCellValue(cell));
                } else {
                    rowData.put(headers.get(i), "");  // Handle empty cells
                }
            }

            data.add(rowData);
        }

        workbook.close();
        fis.close();

        return data;
    }

    // Helper function to get cell values
    public static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }



}
