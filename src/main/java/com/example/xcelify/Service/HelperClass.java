package com.example.xcelify.Service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Map;
@Service
@Slf4j
public class HelperClass {



    public static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                log.info("Найдена колонка '{}'", columnName);
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    public static String getCellStringValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell != null && cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        }
        return "";
    }
        public static Workbook createWorkbook(String[] headers) {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Отчёт");
            createHeaderRow(sheet, headers);
            return workbook;
        }

        public static void createHeaderRow(Sheet sheet, String[] headers) {
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
                sheet.setColumnWidth(i, 4500);
            }
            header.setHeightInPoints(30);
        }

        public static void writeDataToSheet(Sheet sheet, int startRow, Map<String, Object[]> data) {
            int rowNum = startRow;
            for (Map.Entry<String, Object[]> entry : data.entrySet()) {
                Row row = sheet.createRow(rowNum++);
                Object[] rowData = entry.getValue();
                for (int colNum = 0; colNum < rowData.length; colNum++) {
                    Cell cell = row.createCell(colNum);
                    if (rowData[colNum] instanceof String) {
                        cell.setCellValue((String) rowData[colNum]);
                    } else if (rowData[colNum] instanceof Integer) {
                        cell.setCellValue((Integer) rowData[colNum]);
                    } else if (rowData[colNum] instanceof Double) {
                        cell.setCellValue((Double) rowData[colNum]);
                    }
                }
            }
        }

    public static double getCellNumericValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return (cell != null && cell.getCellType() == CellType.NUMERIC) ? cell.getNumericCellValue() : 0.0;
    }

        public static Workbook loadWorkbook(File file) throws IOException {
            try (FileInputStream inputStream = new FileInputStream(file)) {
                return new XSSFWorkbook(inputStream);
            }
        }


        public static void saveWorkbook(Workbook workbook, String outputFilePath) throws IOException {
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
            }
        }
}
