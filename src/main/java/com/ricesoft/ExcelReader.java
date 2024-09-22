package com.ricesoft;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

public class ExcelReader {
    public static void main(String[] args) {
        String excelFilePath = "T:\\Project\\Backend\\File\\nhanh.xlsx";
         try (FileInputStream fis = new FileInputStream(excelFilePath);
                    Workbook workbook = new XSSFWorkbook(fis)) {
             // Get the first sheet
             Sheet sheet = workbook.getSheetAt(0);
             // Read the first row (header row) to find the indices for name of header
             Row headerRow = sheet.getRow(0);
             if(headerRow == null){
                 System.out.println("No header row found.");
                 return;
             }

             Map<String, Integer> columnIndices = new HashMap<>();
             for (Cell cell : headerRow){
                 String headerValue = cell.getStringCellValue();
                 if("SLHĐ (Tổng bán)".equalsIgnoreCase(headerValue)){
                     columnIndices.put("SLHĐ (Tổng bán)", cell.getColumnIndex());
                 }
                 if("SLHĐ (Tổng trả)".equalsIgnoreCase(headerValue)){
                     columnIndices.put("SLHĐ (Tổng trả)", cell.getColumnIndex());
                 }
                 if("Thực thu tiền mặt".equalsIgnoreCase(headerValue)){
                     columnIndices.put("Thực thu tiền mặt", cell.getColumnIndex());
                 }
             }
             if (!columnIndices.containsKey("SLHĐ (Tổng bán)") || !columnIndices.containsKey("SLHĐ (Tổng trả)") || !columnIndices.containsKey("Thực thu tiền mặt")) {
                 System.out.println("Required headers 'Name' or 'Location' not found.");
                 return;
             }

             // Define DecimalFormat to format numbers without scientific notation
             DecimalFormat df = new DecimalFormat("#");

             // Iterate through the remaining rows and print only "SLHĐ (Tổng bán)", "SLHĐ (Tổng trả)", "Thực thu tiền mặt" columns
             System.out.println("SLHĐ (Tổng bán)\tSLHĐ (Tổng trả)\tThực thu tiền mặt");
             for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                 Row row = sheet.getRow(rowIndex);
                 if (row == null) {
                     continue; // Skip empty rows
                 }

                 Cell sldhTotalSellCell = row.getCell(columnIndices.get("SLHĐ (Tổng bán)"));
                 Cell totalPayCell = row.getCell(columnIndices.get("SLHĐ (Tổng trả)"));
                 Cell receiveCell = row.getCell(columnIndices.get("Thực thu tiền mặt"));

                 // Get the cell values (handle possible nulls)
                 String totalSell = getFormattedCellValue(sldhTotalSellCell, df);
                 String totalPay = getFormattedCellValue(totalPayCell, df);
                 String receive = getFormattedCellValue(receiveCell, df);

                 // Print the formatted values
                 System.out.println(totalSell + "\t" + totalPay + "\t" + receive);
             }
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getFormattedCellValue(Cell cell, DecimalFormat df) {
        if (cell == null) {
            return ""; // If the cell is null, return an empty string
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                // Check if the cell contains a date
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Handle date formatting if necessary (not needed in your case, but you can format dates here if needed)
                    return cell.getDateCellValue().toString();
                } else {
                    // Format numeric values as plain decimals (no scientific notation)
                    return df.format(cell.getNumericCellValue());
                }
            case STRING:
                return cell.getStringCellValue(); // If it's a string, return the string value
            case BLANK:
                return ""; // If it's a blank cell, return an empty string
            default:
                return cell.toString(); // For any other cell type (boolean, etc.), return the default string representation
        }

    }
}
