package com.convert.pdf2xls.controller;


import java.util.HashMap;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import com.convert.pdf2xls.service.ConvertApiService;

@RestController
@RequestMapping("/api/pdf")
public class PdfConversionController {

    private final ConvertApiService convertApiService;

    public PdfConversionController(ConvertApiService convertApiService) {
        this.convertApiService = convertApiService;
    }

    @PostMapping("/convert")
    public ResponseEntity<Resource> convertPdfToExcel(@RequestParam("file") MultipartFile file) {
        try {
            File tempFile = convertMultiPartToFile(file);

            Path excelFilePath = convertApiService.convertPdfToExcel(tempFile.toPath());

            // Ensure that the Excel file is downloadable by setting proper headers
            FileSystemResource resource = new FileSystemResource(excelFilePath.toFile());

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + excelFilePath.getFileName());
            headers.add(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE);

            return new ResponseEntity<>(resource, headers, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping("/extract")
    public ResponseEntity<Map<String, Object>> extractDataFromPdf(@RequestParam("file") MultipartFile file) {
        try {
            File tempFile = convertMultiPartToFile(file);

            Path excelFilePath = convertApiService.convertPdfToExcel(tempFile.toPath());
            List<Map<String, String>> extractedData = convertApiService.extractDataFromExcel(excelFilePath);

            Map<String, Object> response = new HashMap<>();
            response.put("data", extractedData);
            response.put("excelFileName", excelFilePath.getFileName().toString());

            return new ResponseEntity<>(response, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    private File convertMultiPartToFile(MultipartFile file) throws IOException {
        File convFile = new File(file.getOriginalFilename());
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(file.getBytes());
        }
        return convFile;
    }
}

