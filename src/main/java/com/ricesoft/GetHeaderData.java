package com.ricesoft;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

public class GetHeaderData {

    public List<Header> getHeader(Row row){
        List<Header> headerList = new ArrayList();

        for(Cell cell : row){
            headerList.add(new Header(cell.getStringCellValue(), cell.getColumnIndex()));
        }

       return headerList;

    }
}
