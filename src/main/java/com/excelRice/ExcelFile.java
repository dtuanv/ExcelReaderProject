package com.excelRice;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;



public class ExcelFile {

    public List<Header> getHeader(String excelFilePath){
        List<Header> headerList = new ArrayList();

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            // Read the first row (header row) to find the indices for name of header
            Row headerRow = sheet.getRow(0);
            if(headerRow == null){
                System.out.println("No header row found.");
                return null;
            }else {
                for(Cell cell : headerRow){
                    headerList.add(new Header(cell.getStringCellValue(), cell.getColumnIndex()));
                }
            }
        }catch (IOException e) {
            e.printStackTrace();
        }

       return headerList;
    }

    public Sheet getSheet(String excelFilePath) {

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            return sheet;
        } catch (IOException e) {
            e.printStackTrace();

        }
        return null;
    }
    public  void extracted(String excelFilePath, String exportFile) {


            ExcelFile getJsonData = new ExcelFile();
            List<Header> headers = getHeader(excelFilePath);

            Set<Integer> columnIndex = new HashSet<>();
            columnIndex.add(1);
            columnIndex.add(2);
            columnIndex.add(6);
            List<Header> headerNew = filterHeader(columnIndex, headers);

        System.out.println("!!!!!!!!!!!!!!!");

            String showHeader = "";
            for(Header h : headerNew){
                showHeader = showHeader + " " + h.name;
            }

            System.out.println(showHeader);
            DecimalFormat df = new DecimalFormat("#");

            Sheet sheet = getSheet(excelFilePath);
        createExcel(sheet,headerNew,exportFile);

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                String showRow = "";
                Row row = sheet.getRow(rowIndex);

                for(Header h : headerNew){
                    showRow = showRow + " | " + getFormattedCellValue(row.getCell(h.index), df);
                }
                System.out.println(showRow);
            }

    }

    public  void createExcel(Sheet sheet, List<Header> selectedHeader, String exportFile){
        DecimalFormat df = new DecimalFormat("#");
        Workbook workbook = new XSSFWorkbook();
        // Create a Sheet
        Sheet sheetNew = workbook.createSheet();
        Row headerRow = sheetNew.createRow(0);
        int cellIndex = 0;
        for (Header h : selectedHeader){
            Cell headerCell = headerRow.createCell(cellIndex);
            headerCell.setCellValue(h.name);
            cellIndex++;
        }
        for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
            Row createdrRow = sheetNew.createRow(rowIndex);
            int cellRowIndex= 0;
            for(Header he : selectedHeader){
                Cell cell = createdrRow.createCell(cellRowIndex);
                cell.setCellValue(getFormattedCellValue(sheet.getRow(rowIndex).getCell(he.index), df));
                cellRowIndex++;
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream(exportFile)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Closing the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Excel file written successfully!");

    }

    public  List<Header> filterHeader(Set<Integer> indexColumn, List<Header> headerList ){
        List<Header>    headerListNew = headerList.stream()
                .filter(header -> indexColumn.stream().anyMatch(integer -> integer == header.index)).collect(Collectors.toList());

        return headerListNew;

    }
    public String getFormattedCellValue(Cell cell, DecimalFormat df) {
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
