package com.convert.pdf2xls.service;

import com.convertapi.client.*;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.CompletableFuture;

@Service
public class ConvertApiService {

    public ConvertApiService() {

        Config.setDefaultSecret("ou8avIN5Qkd7YF3a");
    }

    public Path convertPdfToExcel(Path pdfFilePath) throws Exception {
        try {
            CompletableFuture<ConversionResult> result = ConvertApi.convert("pdf", "xlsx", new Param("file", pdfFilePath));
            String uniqueFileName = "output_" + UUID.randomUUID() + ".xlsx";
            Path outputExcelPath = Paths.get(uniqueFileName);
            result.get().saveFile(outputExcelPath).get();
            return outputExcelPath;
        } catch (ConversionException e) {
            System.err.println("Conversion failed: " + e.getMessage());
            throw e;
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            throw e;
        }

    }

    public List<Map<String, String>> extractDataFromExcel(Path excelFilePath) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();
        double runningBalance = 0.0;

        try (FileInputStream fis = new FileInputStream(excelFilePath.toFile())) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);


            Map<String, Integer> columnMapping = detectRelevantColumns(sheet);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowData = new HashMap<>();


                Cell dateCell = row.getCell(columnMapping.getOrDefault("Date", -1));
                String date = cellToString(dateCell);
                if (date.isEmpty()) {
                    continue;
                }
                rowData.put("Date", date);


                Cell balanceCell = row.getCell(columnMapping.getOrDefault("Balance", -1));
                String balanceString = cellToString(balanceCell).replaceAll("[^\\d.]", "");
                runningBalance = balanceString.isEmpty() ? runningBalance : Double.parseDouble(balanceString);
                rowData.put("Balance", String.valueOf(runningBalance)); // Store balance as a string


                String amount = "";
                String transactionType = "";


                Cell debitCell = row.getCell(columnMapping.getOrDefault("Debit", -1));
                Cell creditCell = row.getCell(columnMapping.getOrDefault("Credit", -1));
                Cell withdrawalCell = row.getCell(columnMapping.getOrDefault("Withdrawal", -1));
                Cell depositCell = row.getCell(columnMapping.getOrDefault("Deposit", -1));

                if (debitCell != null && debitCell.getCellType() == CellType.NUMERIC) {
                    amount = String.valueOf(debitCell.getNumericCellValue());
                    transactionType = "debit";
                    runningBalance -= Double.parseDouble(amount);
                } else if (creditCell != null && creditCell.getCellType() == CellType.NUMERIC) {
                    amount = String.valueOf(creditCell.getNumericCellValue());
                    transactionType = "credit";
                    runningBalance += Double.parseDouble(amount);
                } else if (withdrawalCell != null && withdrawalCell.getCellType() == CellType.NUMERIC) {
                    amount = String.valueOf(withdrawalCell.getNumericCellValue());
                    transactionType = "withdrawal";
                    runningBalance -= Double.parseDouble(amount);
                } else if (depositCell != null && depositCell.getCellType() == CellType.NUMERIC) {
                    amount = String.valueOf(depositCell.getNumericCellValue());
                    transactionType = "deposit";
                    runningBalance += Double.parseDouble(amount);
                }


                rowData.put("Amount", amount);
                rowData.put("Transaction type", transactionType);


                if (!rowData.isEmpty() && !rowData.values().stream().allMatch(String::isEmpty)) {
                    data.add(rowData);
                }
            }
        }

        return data;
    }

    private Map<String, Integer> detectRelevantColumns(Sheet sheet) {
        Map<String, Integer> columnMapping = new HashMap<>();
        int maxRowsToCheck = 5;

        for (int i = 0; i < maxRowsToCheck && i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            for (Cell cell : row) {
                String cellValue = cellToString(cell).toLowerCase();
                if (cellValue.contains("date") && !columnMapping.containsKey("Date")) {
                    columnMapping.put("Date", cell.getColumnIndex());
                } else if (cellValue.contains("balance") && !columnMapping.containsKey("Balance")) {
                    columnMapping.put("Balance", cell.getColumnIndex());
                } else if (cellValue.contains("debit") || cellValue.contains("dr")) {
                    columnMapping.put("Debit", cell.getColumnIndex());
                } else if (cellValue.contains("credit") || cellValue.contains("cr")) {
                    columnMapping.put("Credit", cell.getColumnIndex());
                } else if (cellValue.contains("withdrawal") || cellValue.contains("withd")) {
                    columnMapping.put("Withdrawal", cell.getColumnIndex());
                } else if (cellValue.contains("deposit") || cellValue.contains("dep")) {
                    columnMapping.put("Deposit", cell.getColumnIndex());
                } else if (cellValue.contains("description") || cellValue.contains("particulars")) {
                    columnMapping.put("Description", cell.getColumnIndex());
                }
            }
        }


        if (!columnMapping.containsKey("Date")) columnMapping.put("Date", 0);
        if (!columnMapping.containsKey("Debit")) columnMapping.put("Debit", 3);
        if (!columnMapping.containsKey("Credit")) columnMapping.put("Credit", 4);
        if (!columnMapping.containsKey("Withdrawal")) columnMapping.put("Withdrawal", 5);
        if (!columnMapping.containsKey("Deposit")) columnMapping.put("Deposit", 6);
        if (!columnMapping.containsKey("Description")) columnMapping.put("Description", 2);
        if (!columnMapping.containsKey("Balance")) columnMapping.put("Balance", 7);

        return columnMapping;
    }

    private String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }
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
