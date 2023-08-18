package com.hitss.exceltojson.controller;

import java.util.HashMap;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.hitss.exceltojson.service.ExcelService;

@RestController
@RequestMapping("/file")
public class ExcelController {
    @Autowired
    private ExcelService service;

    @GetMapping
    public ResponseEntity<List<HashMap<String, Object>>> convertExcelToJson(
            @RequestParam String fileLocation)
            throws Exception {
        return new ResponseEntity<>(service.convertExcelToJson(fileLocation),HttpStatus.OK);
    }

}


