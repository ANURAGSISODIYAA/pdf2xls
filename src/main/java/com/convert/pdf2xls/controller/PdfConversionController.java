package com.convert.pdf2xls.controller;

import org.springframework.http.HttpStatus;
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
    public ResponseEntity<?> convertPdfToExcel(@RequestParam("file") MultipartFile file) {
        try {

            File tempFile = convertMultiPartToFile(file);

            Path excelFilePath = convertApiService.convertPdfToExcel(tempFile.toPath());

            List<Map<String, String>> extractedData = convertApiService.extractDataFromExcel(excelFilePath);

            return new ResponseEntity<>(extractedData, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>("Failed to convert and extract data: " + e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
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
