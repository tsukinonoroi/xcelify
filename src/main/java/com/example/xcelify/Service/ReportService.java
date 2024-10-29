package com.example.xcelify.Service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashSet;
import java.util.Set;

@Slf4j
@Service
@RequiredArgsConstructor
public class ReportService {

    public Set<String> parseUniqueProducts(MultipartFile file) throws IOException {
        Set<String> uniqueProducts = new HashSet<>();
        log.info("Метод использован");

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);


            int productNameColumnIndex = -1;
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                if ("Название".equalsIgnoreCase(cell.getStringCellValue().trim())) {
                    productNameColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка");
                    break;
                }
            }

            if (productNameColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Название' не найдена в файле.");
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(productNameColumnIndex);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        uniqueProducts.add(cell.getStringCellValue().trim());
                    }
                }
            }
        }

        return uniqueProducts;
    }
}
