package com.javis.sql_query_generator.controller;

import com.javis.sql_query_generator.services.SqlQuery;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.text.ParseException;
import java.util.Map;

@CrossOrigin("*")
@RestController
public class SqlQueryController {

    @PostMapping("/query-generator")
    public ResponseEntity<Map<String,String>> generateQuery(@RequestParam("file") MultipartFile file) throws IOException, ParseException, IllegalAccessException {
        Map<String, String> queryMap= SqlQuery.getSqlQueryGcExcel(file.getInputStream());
        return new ResponseEntity<>(queryMap, HttpStatus.OK);

    }
}
