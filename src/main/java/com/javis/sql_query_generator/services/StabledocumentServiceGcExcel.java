package com.javis.sql_query_generator.services;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;

@Service
public class StabledocumentServiceGcExcel {
    public static StableDocumentXls readFromWorkBook(Workbook workbook) throws IOException {
        return readFromWorkBook(workbook, null);
    }

    public static StableDocumentXls readFromWorkBook(Workbook workbook, Integer maxRows) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<String>> data = new ArrayList<>();
        Map<Integer, String> headerNameMap = new HashMap<>();
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        int totalRows = worksheet.getRowCount();
        totalRows = Math.min(totalRows,1048576);
        int totalColumns = worksheet.getColumnCount();

        int rowCount = 0;
        int headerCount = 0;

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            List<String> rowData = new ArrayList<>();
            for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {

                Object cell = worksheet.getRange(rowIndex, columnIndex).getValue();

                if (rowCount == 0) {
                    String value = null;
                    if(cell != null ){
                        value = cell.toString().trim();
                    }
                    headers.add(value);
                    headerNameMap.put(columnIndex, value);

                } else {
                    if(cell != null){
                        if(cell instanceof LocalDateTime){
                            rowData.add(cell.toString());
                        }
                        else if(cell instanceof Double){
                            String format = worksheet.getRange(rowIndex, columnIndex).getNumberFormat();
                            String value = new BigDecimal(cell.toString()).toPlainString();
                            if(format.indexOf('.') < 0 && value.indexOf('.') >= 0){
                                //if it's an integer, remove floating point
                                String valAfterDecimal = value.substring(value.indexOf('.')+1);
                                if(Double.valueOf(valAfterDecimal).equals(0.0)){
                                    value = value.substring(0,value.indexOf('.'));
                                }
                                rowData.add(value);
                            }
                            else{
                                rowData.add(value);
                            }
                        }
                        else{
                            rowData.add(cell.toString());
                        }

                    }
                    else{

                            rowData.add("");


                    }
                }
            }
            if (rowCount != 0 && !checkEmptyRow(rowData)) {
                data.add(rowData);
            }

            rowCount++;
            if (maxRows != null && rowCount == maxRows)
                break;
        }

        StableDocumentXls xls = StableDocumentXls.builder()
                .headers(headers)
                .data(data)
                .build();
        return xls;
    }

    public static StableDocumentXls readFromWorkBook(Workbook workbook, Integer maxRows, IWorksheet worksheet) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<String>> data = new ArrayList<>();
        Map<String, Integer> columnIndexMap = new HashMap<>();
        Map<Integer, String> indexColumnMap = new HashMap<>();
        int totalRows = worksheet.getRowCount();
        totalRows = Math.min(totalRows,1048576);
        int totalColumns = worksheet.getColumnCount();

        int rowCount = 0;
        int headerCount = 0;

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            List<String> rowData = new ArrayList<>();
            for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {

                Object cell = worksheet.getRange(rowIndex, columnIndex).getValue();

                if (rowCount == 0) {
                    String value = null;
                    if(cell != null ){
                        value = cell.toString().trim();

                    }
                    headers.add(value);
                    columnIndexMap.put(value,columnIndex);
                    indexColumnMap.put(columnIndex, value);

                } else {
                    if(cell != null){
                        if(cell instanceof LocalDateTime){
                            rowData.add(cell.toString());
                        }
                        else if(cell instanceof Double){
                            String format = worksheet.getRange(rowIndex, columnIndex).getNumberFormat();
                            String value = new BigDecimal(cell.toString()).toPlainString();
                            if(format.indexOf('.') < 0 && value.indexOf('.') >= 0){
                                //if it's an integer, remove floating point
                                String valAfterDecimal = value.substring(value.indexOf('.')+1);
                                if(Double.valueOf(valAfterDecimal).equals(0.0)){
                                    value = value.substring(0,value.indexOf('.'));
                                }
                                rowData.add(value);
                            }
                            else{
                                rowData.add(value);
                            }
                        }
                        else{
                            rowData.add(cell.toString());
                        }

                    }
                    else{

                        rowData.add("");


                    }
                }
            }
            if (rowCount != 0 && !checkEmptyRow(rowData)) {
                data.add(rowData);
            }

            rowCount++;
            if (maxRows != null && rowCount == maxRows)
                break;
        }

        StableDocumentXls xls = StableDocumentXls.builder()
                .headers(headers)
                .data(data)
                .indexColumnMap(indexColumnMap)
                .columnIndexMap(columnIndexMap)
                .build();
        return xls;
    }

    private static boolean checkEmptyRow(List<String> rowData) {
        for (String data : rowData) {
            if (!data.trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }
    public void setColumnMap(StableDocumentXls documentXls){
        Map<String, Integer> columnIndexMap = new HashMap<>();
        Map<Integer, String> indexColumnMap = new HashMap<>();
        int position = 0;
        for(String header: documentXls.getHeaders()){
            columnIndexMap.put(header, position);
            indexColumnMap.put(position, header);
            position++;
        }
        documentXls.setColumnIndexMap(columnIndexMap);
        documentXls.setIndexColumnMap(indexColumnMap);
    }

    private static void insertRow(IWorksheet firstSheet, List<String> row,int colSize, int rowNum) {
        int cellNum = 0;

        for (String obj: row) {
            if(cellNum >= colSize){
                break;
            }
            if(Objects.equals(obj, "TRUE()")||Objects.equals(obj, "true()")||Objects.equals(obj, "True()") )
                firstSheet.getRange(rowNum, cellNum).setValue(true);
            else if(Objects.equals(obj,"FALSE()")||Objects.equals(obj, "false()")||Objects.equals(obj, "False()"))
                firstSheet.getRange(rowNum, cellNum).setValue(false);
            else
                firstSheet.getRange(rowNum, cellNum).setValue(obj);
            cellNum++;
        }

    }
    public static Workbook writeToWorkBook(StableDocumentXls xls) {
        Workbook workbook = new Workbook();
        IWorksheet firstSheet = workbook.getWorksheets().get(0);
        insertRow(firstSheet, xls.getHeaders(), xls.getHeaders().size(), 0);
        int rowNum = 1;
        if (xls.getData().size() > 0) {
            for (List<String> rowData : xls.getData()) {
                if(rowNum >= 1048576){
                    break;
                }
                insertRow(firstSheet, rowData,xls.getHeaders().size(), rowNum);
                rowNum++;
            }
        }
        return workbook;
    }

    public Workbook listsToWorkbook(List<List<String>> companyData){
        Integer listSize = companyData.size();
        List<String> headers = companyData.get(listSize - 1);
        companyData.remove(listSize - 1);
        StableDocumentXls xls = StableDocumentXls.builder().build();
        xls.setHeaders(headers);
        xls.setData(companyData);
        Workbook workbook = writeToWorkBook(xls);
        return workbook;
    }

    public void subDocumentXls1(StableDocumentXls xls, int beginIndex, int endIndex) {
        if (xls.getData().isEmpty() ||beginIndex < 0 || endIndex > xls.getData().size())
            return;

        List<List<String>> newData = new ArrayList<>();

        for (int rowIndex = beginIndex; rowIndex < endIndex; rowIndex++) {
            List<String> rowData = xls.getData().get(rowIndex);
            newData.add(rowData);
        }

        xls.setData(newData);
    }

    public int createColumn(StableDocumentXls stableDocumentXls, String columnName) {
        int index = stableDocumentXls.getHeaders().indexOf(columnName);
        if (index == -1) {
            index = stableDocumentXls.getHeaders().size();
            stableDocumentXls.getHeaders().add(index, columnName);
            stableDocumentXls.getIndexColumnMap().put(index,columnName);
            stableDocumentXls.getColumnIndexMap().put(columnName, index);

        }

        if (stableDocumentXls.getData() == null)
            stableDocumentXls.setData(new ArrayList<>());
        for (int rowIndex = 0; rowIndex < stableDocumentXls.getData().size(); rowIndex++) {
            List<String> row = stableDocumentXls.getData().get(rowIndex);
//			Map<String, String> rowByHeaderName = stableDocumentXls.getDataByHeaderName().get(rowIndex);
            for (int left = row.size(); left <= index; left++)
                row.add(left, "");

        }
        return index;
    }

}
