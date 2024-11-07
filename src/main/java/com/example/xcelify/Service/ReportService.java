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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import java.time.LocalDateTime;
import java.util.HashSet;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class ReportService {

    private final ReportRepository reportRepository;
    private final ProductRepository productRepository;

    @Transactional
    public Set<Product> parseUniqueProducts(MultipartFile file) throws IOException {
        long startTime = System.currentTimeMillis(); // Начало отсчёта времени
        Set<Product> uniqueProducts = new HashSet<>();
        log.info("Метод parseUniqueProducts использован");

        // Загрузка всех продуктов один раз
        Map<String, Product> existingProducts = productRepository.findAll().stream()
                .collect(Collectors.toMap(Product::getArticul, Function.identity()));

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            int productNameColumnIndex = -1;
            int supplierArticulColumnIndex = -1;
            Row headerRow = sheet.getRow(0);

            // Находим индексы колонок
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

            // Обработка строк
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

    public void generateAndSaveReport(MultipartFile file) throws IOException {
        Set<Product> uniqueProducts = parseUniqueProducts(file);


        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Итоговый отчет");

        Row resultHeader = resultSheet.createRow(0);
        String[] columns = {"Название", "Артикул поставщика", "Кол-во",
                "Вайлдберриз реализовал Товар (Пр)", "Возмещение за выдачу и возврат товаров на ПВЗ",
                "Эквайринг/Комиссия за организацию платежей", "Вознаграждение Вайлдберриз (ВВ), без НДС",
                "НДС с Вознаграждения Вайлдберриз", "К перечислению Продавцу за реализованный Товар",
                "Услуги по доставке товара покупателю", "Общая сумма штрафов", "Хранение",
                "Удержания", "Платная приемка", "к оплате", "себестоимость", "Прибыль до налогообложения"};
        for (int j = 0; j < columns.length; j++) {
            resultHeader.createCell(j).setCellValue(columns[j]);
        }


        int rowIndex = 1;
        for (Product product : uniqueProducts) {
            Row row = resultSheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(product.getName());
            row.createCell(1).setCellValue(product.getArticul());
            row.createCell(2).setCellValue(1);


            for (int col = 3; col < columns.length; col++) {
                row.createCell(col).setCellValue(0);
            }
        }


        String directoryPath = "src/main/java/com/example/xcelify/Reports/Filter";
        File directory = new File(directoryPath);
        if (!directory.exists()) {
            directory.mkdirs();
        }
        
        try (FileOutputStream fileOut = new FileOutputStream(new File(directoryPath, "final_report.xlsx"))) {
            resultWorkbook.write(fileOut);
        }

        System.out.println("Отчет успешно сохранен в " + directoryPath + "/final_report.xlsx");
    }
}

