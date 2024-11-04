package com.example.xcelify.Service;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Repository.ReportRepository;
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
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;

@Slf4j
@Service
@RequiredArgsConstructor
public class ReportService {

    private final ReportRepository reportRepository;
    private final ProductRepository productRepository;

    public Set<Product> parseUniqueProducts(MultipartFile file) throws IOException {
        Set<Product> uniqueProducts = new HashSet<>();
        log.info("Метод parseUniqueProducts использован");

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            int productNameColumnIndex = -1;
            int supplierArticulColumnIndex = -1;
            Row headerRow = sheet.getRow(0);

            for (Cell cell : headerRow) {
                String cellValue = cell.getStringCellValue().trim();
                if ("Название".equalsIgnoreCase(cellValue)) {
                    productNameColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка 'Название'");
                } else if ("Артикул поставщика".equalsIgnoreCase(cellValue)) {
                    supplierArticulColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка 'Артикул поставщика'");
                }
            }

            if (productNameColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Название' не найдена в файле.");
            }

            if (supplierArticulColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Артикул поставщика' не найдена в файле.");
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell nameCell = row.getCell(productNameColumnIndex);
                    Cell articulCell = row.getCell(supplierArticulColumnIndex);

                    String name = (nameCell != null && nameCell.getCellType() == CellType.STRING)
                            ? nameCell.getStringCellValue().trim()
                            : "";
                    String articul = (articulCell != null && articulCell.getCellType() == CellType.STRING)
                            ? articulCell.getStringCellValue().trim()
                            : "";

                    if (!name.isEmpty() && !articul.isEmpty()) {
                        Product product = new Product();
                        product.setName(name);
                        product.setArticul(articul);
                        uniqueProducts.add(product);
                    } else {
                        log.warn("Пропущена строка: Название = '{}', Артикул = '{}'", name, articul);
                    }
                }
            }
        }

        productRepository.saveAll(uniqueProducts);
        return uniqueProducts;
    }
}
