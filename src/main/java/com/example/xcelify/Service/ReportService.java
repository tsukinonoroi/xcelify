package com.example.xcelify.Service;
import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Repository.ReportRepository;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;


import java.io.*;

import java.time.LocalDateTime;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
@Data
public class ReportService {

    private final ReportRepository reportRepository;
    private final ProductRepository productRepository;
    private final HelperClass helperClass;
    private String reportName;
    private String filePath = "C:/users/edemw/Desktop/filter";
    private String sourceFilePath;

    @Transactional
    public Set<Product> parseUniqueProducts(File file) throws IOException {
        long startTime = System.currentTimeMillis(); // Начало отсчёта времени
        Set<Product> uniqueProducts = new HashSet<>();
        log.info("Метод parseUniqueProducts использован");

        Map<String, Product> existingProducts = productRepository.findAll().stream()
                .collect(Collectors.toMap(Product::getArticul, Function.identity()));

        try (Workbook workbook = HelperClass.loadWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            int productNameColumnIndex = HelperClass.findColumnIndex(sheet, "Название");
            int supplierArticulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");

            if (productNameColumnIndex == -1 || supplierArticulColumnIndex == -1) {
                throw new IllegalArgumentException("Не найдены необходимые колонки в файле.");
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    String name = HelperClass.getCellStringValue(row, productNameColumnIndex);
                    String articul = HelperClass.getCellStringValue(row, supplierArticulColumnIndex);

                    if (!name.isEmpty() && !articul.isEmpty()) {
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
                    }
                }
            }
        }

        long endTime = System.currentTimeMillis();
        log.info("Время выполнения метода parseUniqueProducts: {} ms", (endTime - startTime));

        return uniqueProducts;
    }

    public void generateNewReport(String reportName) throws IOException {
        String uniqueFileName = reportName + ".xlsx";
        File outputDir = new File(filePath);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        String outputFilePath = outputDir.getAbsolutePath() + "/" + uniqueFileName;
        String[] headers = {
                "Название", "Артикул поставщика", "Кол-во",
                "Вайлдберриз реализовал Товар (Пр)", "Возмещение за выдачу и возврат товаров на ПВЗ",
                "Эквайринг/Комиссия за организацию платежей", "Вознаграждение Вайлдберриз (ВВ), без НДС",
                "НДС с Вознаграждения Вайлдберриз", "К перечислению Продавцу за реализованный Товар",
                "Услуги по доставке товара покупателю", "Общая сумма штрафов", "Хранение", "Удержания",
                "Платная приемка", "к оплате", "себестоимость", "Прибыль до налогообложения"
        };

        Workbook workbook = HelperClass.createWorkbook(headers);
        Sheet sheet = workbook.getSheetAt(0);

        // Устанавливаем высоту строк на 30 пунктов
        int rowHeightInPoints = 30;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            row.setHeightInPoints(rowHeightInPoints);
        }

        HelperClass.saveWorkbook(workbook, outputFilePath);

        log.info("Отчет успешно создан: " + outputFilePath);
    }


    public void inputNameAndArticul(Set<Product> uniqueProducts) throws IOException {
        String outputFilePath = filePath + "/" + reportName + ".xlsx";
        File file = new File(outputFilePath);

        if (!file.exists()) {
            throw new FileNotFoundException("Файл отчета не найден: " + outputFilePath);
        }

        Workbook workbook = HelperClass.loadWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, Object[]> data = uniqueProducts.stream().collect(Collectors.toMap(
                Product::getArticul,
                product -> new Object[]{product.getName(), product.getArticul()}
        ));
        HelperClass.writeDataToSheet(sheet, 1, data);
        HelperClass.saveWorkbook(workbook, outputFilePath);

        log.info("Отчет успешно обновлен с данными продуктов: " + outputFilePath);
    }

    public void inputCountAndSale(Set<Product> uniqueProducts) throws IOException {
        if (sourceFilePath == null) {
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файл перед генерацией отчета.");
        }
        Map<String, Integer> countMap = new HashMap<>();
        Map<String, Double> saleSumMap = new HashMap<>();

        try (Workbook workbook = HelperClass.loadWorkbook(new File(sourceFilePath))) {
            Sheet sheet = workbook.getSheetAt(0);

            int nameColumnIndex = HelperClass.findColumnIndex(sheet, "Название");
            int supplierArticulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");
            int saleAmountColumnIndex = HelperClass.findColumnIndex(sheet, "Вайлдберриз реализовал Товар (Пр)");

            if (supplierArticulColumnIndex == -1 || saleAmountColumnIndex == -1) {
                throw new IllegalArgumentException("Не найдены необходимые колонки в исходном файле.");
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    String articul = HelperClass.getCellStringValue(row, supplierArticulColumnIndex);
                    double saleAmount = row.getCell(saleAmountColumnIndex).getNumericCellValue();

                    if (uniqueProducts.stream().anyMatch(product -> product.getArticul().equals(articul))) {
                        countMap.put(articul, countMap.getOrDefault(articul, 0) + (saleAmount != 0 ? 1 : 0));
                        saleSumMap.put(articul, saleSumMap.getOrDefault(articul, 0.0) + saleAmount);
                    }
                }
            }
        }

        String outputFilePath = filePath + "/" + reportName + ".xlsx";
        Workbook reportWorkbook;

        File file = new File(outputFilePath);
        if (file.exists()) {
            reportWorkbook = HelperClass.loadWorkbook(file);
        } else {
            String[] headers = {"Название", "Артикул поставщика", "Кол-во", "Вайлдберриз реализовал товар (Пр)"};
            reportWorkbook = HelperClass.createWorkbook(headers);
        }

        Sheet reportSheet = reportWorkbook.getSheetAt(0);
        if (reportSheet == null) {
            reportSheet = reportWorkbook.createSheet("Отчет");
        }

        Map<String, Object[]> reportData = new HashMap<>();
        int rowNum = 1;
        for (Product product : uniqueProducts) {
            Object[] rowData = {
                    product.getName(),
                    product.getArticul(),
                    countMap.getOrDefault(product.getArticul(), 0),
                    saleSumMap.getOrDefault(product.getArticul(), 0.0)
            };
            reportData.put(String.valueOf(rowNum++), rowData);
        }

        HelperClass.writeDataToSheet(reportSheet, 1, reportData);
        HelperClass.saveWorkbook(reportWorkbook, outputFilePath);
        System.out.println("Отчет успешно обновлен с количеством и суммой продаж: " + outputFilePath);
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
    
    public File getSourceFile() {
        if (sourceFilePath == null) {
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файл перед генерацией отчета.");
        }
        return new File(sourceFilePath);
    }
}

