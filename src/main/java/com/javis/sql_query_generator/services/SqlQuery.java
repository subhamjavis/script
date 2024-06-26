package com.javis.sql_query_generator.services;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.util.*;
import java.util.stream.Collectors;


public class SqlQuery {

    public static Map<String, String> getSqlQuery(InputStream inputStream) throws IOException, IllegalAccessException, ParseException {
        Map<String, String> queryMap = new HashMap<>();
        ReadableWorkbook readableWorkbook = new ReadableWorkbook(inputStream);
        for(Sheet sheet : readableWorkbook.getSheets().collect(Collectors.toList())){
            StableDocumentXls documentXls = StableDocumentServiceFastExcel.readFromWorkBook(readableWorkbook, null, sheet);
            StringBuilder query = SqlQuery.generateUpdateSqlQuery(documentXls, sheet.getName());
            if(query!=null){
                queryMap.put(sheet.getName(), query.toString());
            }

        }
        inputStream.close();
        return queryMap;
    }

    public static Map<String, String> getSqlQueryGcExcel(InputStream inputStream) throws IOException, IllegalAccessException, ParseException {
        Map<String, String> queryMap = new HashMap<>();
        Workbook workbook = new Workbook();
        workbook.open(inputStream);
        for(IWorksheet worksheet : workbook.getWorksheets()){
            StableDocumentXls documentXls = StabledocumentServiceGcExcel.readFromWorkBook(workbook, null, worksheet);
            StringBuilder updatequery = SqlQuery.generateUpdateSqlQuery(documentXls, worksheet.getName());
//            StringBuilder insertquery = SqlQuery.generateInsertSqlQuery(documentXls, worksheet.getName());
            if(updatequery!=null){
                queryMap.put(worksheet.getName(), updatequery.toString());
            }
//            if(insertquery!=null){
//                queryMap.put(worksheet.getName()+"~insert", insertquery.toString());
//            }

        }
        inputStream.close();
        return queryMap;
    }

//    public static void main(String[] args) throws IOException, IllegalAccessException, ParseException {
//        Map<String, String> queryMap = new HashMap<>();
//        File file = new File("/Users/subham/MyProjects/demo/alternate-5.xlsx");
//        ReadableWorkbook readableWorkbook = new ReadableWorkbook(file);
//        for(Sheet sheet : readableWorkbook.getSheets().collect(Collectors.toList())){
//            StableDocumentXls documentXls = StableDocumentServiceFastExcel.readFromWorkBook(readableWorkbook, null, sheet);
//            StringBuilder query = SqlQuery.generateUpdateSqlQuery(documentXls, sheet.getName());
//            queryMap.put(sheet.getName(), query.toString());
//        }
//
//
//    }
    public static <T>StringBuilder generateUpdateSqlQuery(StableDocumentXls documentXls, String tablename) throws IllegalAccessException {
        if(documentXls.getData().size()==0){
            return null;
        }
        StringBuilder update = generateInsertSqlQuery(documentXls, tablename);
        update.append("ON DUPLICATE KEY UPDATE ");

        int c=0;
        for(String column : documentXls.getHeaders()){
            if(column != null && !column.isEmpty()){

                if(c==0){
                    update.append(column+" = values ("+column+")");

                }
                else{

                    update.append(","+column+" = values ("+column+")");
                }

            }
            c++;
        }


        return update;
    }

    public static <T>StringBuilder generateInsertSqlQuery(StableDocumentXls documentXls, String tablename) throws IllegalAccessException {
        StringBuilder insert = new StringBuilder("Insert into ");
        insert.append(tablename+" ");

        StringBuilder objectQuery = new StringBuilder();
        int c=0;
        for(String columnName:documentXls.getHeaders()){

            if(columnName != null && !columnName.isEmpty()){
                if(c==0){

                    objectQuery.append(columnName);

                }
                else{

                    objectQuery.append(","+columnName);


                }

            }
            c++;
        }
        if(objectQuery.length()>0){
            insert.append("("+objectQuery+")");
        }

        if(documentXls.getData().size()>0){
            insert.append(" values ");
        }
        for(int i=0;i<documentXls.getData().size();i++){
            if(i==0){
                insert.append("(");
            }
            else{
                insert.append(",(");
            }
            c=0;
            for(String row : documentXls.getData().get(i)){
                 String column = documentXls.getIndexColumnMap().get(c);
                if(column != null && !column.isEmpty()){
                    if(c==0){

                            if(row==null || row.isEmpty()){
                                if(row.isEmpty()){
                                    row=null;
                                }
                                insert.append(row);
                            }
                            else{
                                insert.append("\'"+row.replace("'", "''")+"\'");
                            }
                    }
                    else{
                            if(row==null || row.isEmpty()){
                                if(row.isEmpty()){
                                    row=null;
                                }
                                insert.append(","+row);
                            }
                            else{
                                insert.append(",\'"+row.replace("'", "''")+"\'");
                            }

                    }

                }
                c++;
            }
            insert.append(")");

        }
        return insert;
    }


}
