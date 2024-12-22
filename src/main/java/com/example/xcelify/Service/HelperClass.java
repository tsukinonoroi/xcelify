package com.example.xcelify.Service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.Map;
@Service
@Slf4j
public class HelperClass {

    // Метод для извлечения и преобразования даты в LocalDate
    public static LocalDate getDateCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue().toLocalDate();
                    } else {
                        // Если ячейка числовая, но не отформатирована как дата,
                        // попробуем интерпретировать её как серийный номер даты Excel
                        return LocalDate.of(1900, 1, 1).plusDays((long) cell.getNumericCellValue() - 2);
                    }
                case STRING:
                    String dateString = cell.getStringCellValue().trim();
                    // Попробуем несколько форматов дат
                    DateTimeFormatter[] formatters = {
                            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                            DateTimeFormatter.ofPattern("dd.MM.yyyy"),
                            DateTimeFormatter.ofPattern("yyyy.MM.dd"),
                            DateTimeFormatter.ofPattern("dd/MM/yyyy")
                    };
                    for (DateTimeFormatter formatter : formatters) {
                        try {
                            return LocalDate.parse(dateString, formatter);
                        } catch (DateTimeParseException e) {
                            // Продолжаем со следующим форматом
                        }
                    }
                    log.warn("Не удалось распознать дату из строки: {}", dateString);
                    break;
                default:
                    log.warn("Неподдерживаемый тип ячейки: {}", cell.getCellType());
            }
        } catch (Exception e) {
            log.error("Не удалось преобразовать ячейку в дату: {}", cell, e);
        }
        return null;
    }


    public static String getSalesPeriod(Sheet sheet) {
        // Поиск строки "Период продаж"
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue().trim();
                        if (cellValue.startsWith("Период продаж:")) {
                            // Извлечение периода продаж из строки
                            String period = cellValue.replace("Период продаж:", "").trim();
                            if (period.matches("\\d{4}-\\d{2}-\\d{2} - \\d{4}-\\d{2}-\\d{2}")) {
                                return period;  // Возвращаем период продаж
                            }
                        }
                    }
                }
            }
        }
        log.error("Не удалось найти строку 'Период продаж' в листе");
        return null;  // Если период не найден
    }

    public static double getNumericCellValueOrDefault(Cell cell, double defaultValue) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return defaultValue; // Возвращаем значение по умолчанию
        }
        try {
            return cell.getNumericCellValue();
        } catch (IllegalStateException e) {
            // Если тип ячейки не числовой, возвращаем значение по умолчанию
            return defaultValue;
        }
    }


    public static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    public static Workbook loadWorkbook(MultipartFile file) throws IOException { try (InputStream inputStream = file.getInputStream()) { return new XSSFWorkbook(inputStream); } }
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
            header.setHeightInPoints(100);
        }

    public static Row findRowByArticul(Sheet sheet, String articul) {
        int articulColumnIndex = findColumnIndex(sheet, "Артикул поставщика");
        if (articulColumnIndex == -1) {
            throw new IllegalArgumentException("Не найдена колонка 'Артикул поставщика' в итоговом отчете.");
        }

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                String cellArticul = getCellStringValue(row, articulColumnIndex);
                if (cellArticul != null && cellArticul.equals(articul)) {
                    return row;
                }
            }
        }
        return null;
    }

    public static void setNumericCellValue(Row row, int cellIndex, double value) {
        Cell cell = row.getCell(cellIndex);
        if (cell == null) {
            cell = row.createCell(cellIndex);
        }
        cell.setCellValue(value);
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

    public static double getNumericCellValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue());
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
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
