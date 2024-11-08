package com.example.xcelify.Service;
import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Repository.ReportRepository;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.server.ResponseStatusException;

import java.io.*;

import java.time.LocalDateTime;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class ReportService {

    private final ReportRepository reportRepository;
    private final ProductRepository productRepository;

    @Transactional
    public Set<Product> parseUniqueProducts(File file) throws IOException {
        long startTime = System.currentTimeMillis(); // Начало отсчёта времени
        Set<Product> uniqueProducts = new HashSet<>();
        log.info("Метод parseUniqueProducts использован");

        
        Map<String, Product> existingProducts = productRepository.findAll().stream()
                .collect(Collectors.toMap(Product::getArticul, Function.identity()));


        Map<String, Integer> soldProducts = new HashMap<>();

        try (FileInputStream inputStream = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            int productNameColumnIndex = -1;
            int supplierArticulColumnIndex = -1;
            int paymentReasonColumnIndex = -1;
            Row headerRow = sheet.getRow(0);


            for (Cell cell : headerRow) {
                String cellValue = cell.getStringCellValue().trim();
                if ("Название".equalsIgnoreCase(cellValue)) {
                    productNameColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка 'Название'");
                } else if ("Артикул поставщика".equalsIgnoreCase(cellValue)) {
                    supplierArticulColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка 'Артикул поставщика'");
                } else if ("Обоснование для оплаты".equalsIgnoreCase(cellValue)) {
                    paymentReasonColumnIndex = cell.getColumnIndex();
                    log.info("Найдена колонка 'Обоснование для оплаты'");
                }
            }

            if (productNameColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Название' не найдена в файле.");
            }

            if (supplierArticulColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Артикул поставщика' не найдена в файле.");
            }

            if (paymentReasonColumnIndex == -1) {
                throw new IllegalArgumentException("Колонка 'Обоснование для оплаты' не найдена в файле.");
            }


            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell nameCell = row.getCell(productNameColumnIndex);
                    Cell articulCell = row.getCell(supplierArticulColumnIndex);
                    Cell paymentReasonCell = row.getCell(paymentReasonColumnIndex);

                    String name = (nameCell != null && nameCell.getCellType() == CellType.STRING)
                            ? nameCell.getStringCellValue().trim()
                            : "";
                    String articul = (articulCell != null && articulCell.getCellType() == CellType.STRING)
                            ? articulCell.getStringCellValue().trim()
                            : "";
                    String paymentReason = (paymentReasonCell != null && paymentReasonCell.getCellType() == CellType.STRING)
                            ? paymentReasonCell.getStringCellValue().trim()
                            : "";

                    if (!name.isEmpty() && !articul.isEmpty()) {
                        if ("продажа".equalsIgnoreCase(paymentReason)) {
                            String key = name + "_" + articul;
                            soldProducts.put(key, soldProducts.getOrDefault(key, 0) + 1);
                        }

                        Product existingProduct = existingProducts.get(articul);
                        if (existingProduct != null) {
                            uniqueProducts.add(existingProduct);
                        } else {
                            Product newProduct = new Product();
                            newProduct.setName(name);
                            newProduct.setArticul(articul);
                            productRepository.save(newProduct);
                            uniqueProducts.add(newProduct);
                            existingProducts.put(articul, newProduct);
                        }
                    } else {
                        log.warn("Пропущена строка: Название = '{}', Артикул = '{}'", name, articul);
                    }
                }
            }

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Отчёт");

            Row header = newSheet.createRow(0);
            String[] headers = {"Название", "Артикул поставщика", "Количество продано"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }

            int rowIndex = 1;
            for (Product product : uniqueProducts) {
                Row row = newSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(product.getName());
                row.createCell(1).setCellValue(product.getArticul());

                String key = product.getName() + "_" + product.getArticul();
                Integer soldQuantity = soldProducts.getOrDefault(key, 0);
                row.createCell(2).setCellValue(soldQuantity);
            }

            File outputDir = new File("C:\\Users\\edemw\\Desktop\\filter");
            if (!outputDir.exists()) {
                outputDir.mkdirs();
            }

            String outputFilePath = outputDir.getAbsolutePath() + "\\" + "report_" + System.currentTimeMillis() + ".xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(fileOut);
            }
            newWorkbook.close();
        }

        long endTime = System.currentTimeMillis(); // Конец отсчёта времени
        log.info("Время выполнения метода parseUniqueProducts: {} ms", (endTime - startTime));

        return uniqueProducts;
    }


    @Transactional
    public void updateCosts(Map<Long, Double> costsMap) {
        for (Map.Entry<Long, Double> entry : costsMap.entrySet()) {
            Long productId = entry.getKey();
            Double newCost = entry.getValue();

            Optional<Product> optionalProduct = productRepository.findById(productId);
            if (optionalProduct.isPresent()) {
                Product product = optionalProduct.get();
                product.setCost(newCost);
                product.setUpdateCost(LocalDateTime.now());
                productRepository.save(product);
            } else {
                log.warn("Продукт не найден с ID: {}", productId);
            }
        }
    }


}

