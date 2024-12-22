package com.example.xcelify.Service;
import com.example.xcelify.Model.Product;
import com.example.xcelify.Model.Report;
import com.example.xcelify.Model.User;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Repository.ReportRepository;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;


import java.io.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
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
    private final CustomUserDetailService customUserDetailService;
    private String reportName;
    private String filePath = "C:/users/edemw/Desktop/filter";
    private String sourceRussianFilePath;
    private String sourceInternationalFilePath;


    //парсинг продуктов исходных файлов
    @Transactional
    public Set<Product> parseUniqueProducts(File fileRussia, File fileInternational) throws IOException {
        long startTime = System.currentTimeMillis();
        Set<Product> uniqueProducts = new HashSet<>();
        log.debug("Метод parseUniqueProducts используется для обработки двух файлов");

        User currentUser = customUserDetailService.getCurrentUser();

        Map<String, Product> existingProducts = productRepository.findAllByUser(currentUser).stream()
                .collect(Collectors.toMap(
                        product -> product.getArticul().toLowerCase(),
                        Function.identity()
                ));

        List<File> files = List.of(fileRussia, fileInternational);

        for (File file : files) {
            try (Workbook workbook = HelperClass.loadWorkbook(file)) {
                Sheet sheet = workbook.getSheetAt(0);

                int productNameColumnIndex = HelperClass.findColumnIndex(sheet, "Название");
                int supplierArticulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");

                if (productNameColumnIndex == -1 || supplierArticulColumnIndex == -1) {
                    throw new IllegalArgumentException("Не найдены необходимые колонки в файле: " + file.getName());
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        String name = HelperClass.getCellStringValue(row, productNameColumnIndex);
                        String articul = HelperClass.getCellStringValue(row, supplierArticulColumnIndex);

                        if (!name.isEmpty() && !articul.isEmpty()) {
                            articul = articul.toLowerCase();

                            Product existingProduct = existingProducts.get(articul);
                            if (existingProduct != null) {
                                uniqueProducts.add(existingProduct);
                            } else {
                                Product newProduct = new Product();
                                newProduct.setName(name);
                                newProduct.setArticul(articul);
                                newProduct.setUser(currentUser);
                                productRepository.save(newProduct);
                                uniqueProducts.add(newProduct);
                                existingProducts.put(articul, newProduct);
                            }
                        }
                    }
                }
            }
        }

        long endTime = System.currentTimeMillis();
        log.info("Время выполнения метода parseUniqueProducts: {} мс", (endTime - startTime));

        List<Product> sortedList = uniqueProducts.stream()
                .sorted(Comparator.comparing(Product::getName))
                .collect(Collectors.toList());

        return new LinkedHashSet<>(sortedList); // Преобразование обратно в Set для уникальности и порядка
    }



    public String getOutputReportPath() {
        return filePath + "/" + reportName + ".xlsx";
    }

    //создание пустого отчета
    public void generateNewReport(String reportName) throws IOException {
        File outputDir = new File(filePath);
        User currentUser = customUserDetailService.getCurrentUser();

        Report report = new Report();
        report.setFilePath(getOutputReportPath());
        report.setFilename(reportName);
        report.setUser(currentUser);
        report.setLocalDateTime(LocalDateTime.now());
        reportRepository.save(report);

        if (!outputDir.exists()) {
            outputDir.mkdirs();
            log.info("Родительской папки не существует, создаем папку");
        }

        String outputFilePath = getOutputReportPath();
        log.info("Путь к файлу для сохранения: " + outputFilePath);

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

        int totalColumns = headers.length;
        for (int i = 0; i < totalColumns; i++) {
            sheet.setColumnWidth(i, 8000);
        }

        int rowHeightInPoints = 30;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            row.setHeightInPoints(rowHeightInPoints);
        }

        HelperClass.saveWorkbook(workbook, outputFilePath);
        log.info("Новый пустой отчет успешно создан: " + outputFilePath);
    }


    //записываем название и артикул продукта
    public void inputNameAndArticul(Set<Product> uniqueProducts) throws IOException {
        String outputFilePath = getOutputReportPath();
        File file = new File(outputFilePath);

        if (!file.exists()) {
            log.warn("Файл отчёта не найден");
            throw new FileNotFoundException("Файл отчета не найден: " + outputFilePath);
        }

        Workbook workbook = HelperClass.loadWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, Object[]> data = uniqueProducts.stream().collect(Collectors.toMap(
                Product::getArticul,
                product -> new Object[]{product.getName(), product.getArticul()}
        ));
        HelperClass.writeDataToSheet(sheet, 1, data);
        HelperClass.saveWorkbook(workbook, getOutputReportPath());

        log.info("Отчет успешно обновлен с данными продуктов: " + getOutputReportPath());
    }

    //считаем/записываем количество продаж и сумму продаж товара
    public synchronized void inputCountAndSale(Set<Product> uniqueProducts) throws IOException {
        if (uniqueProducts == null || uniqueProducts.isEmpty()) {
            throw new IllegalStateException("Уникальные продукты не переданы или пусты.");
        }

        Map<String, Integer> countMap = new HashMap<>();
        Map<String, Double> saleSumMap = new HashMap<>();
        Map<String, Integer> returnCountMap = new HashMap<>();
        Map<String, Double> returnSumMap = new HashMap<>();

        // Список файлов для обработки
        File russianFile = new File(sourceRussianFilePath);
        File internationalFile = new File(sourceInternationalFilePath);

        List<File> files = List.of(russianFile, internationalFile);

        for (File file : files) {
            try (Workbook workbook = HelperClass.loadWorkbook(file)) {
                log.info("Пытаемся открыть отчет...");
                Sheet sheet = workbook.getSheetAt(0);

                int supplierArticulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");
                int saleAmountColumnIndex = HelperClass.findColumnIndex(sheet, "Вайлдберриз реализовал Товар (Пр)");
                int reasonForPaymentColumnIndex = HelperClass.findColumnIndex(sheet, "Обоснование для оплаты");
                int quantityColumnIndex = HelperClass.findColumnIndex(sheet, "Кол-во");

                if (supplierArticulColumnIndex == -1 || saleAmountColumnIndex == -1 || reasonForPaymentColumnIndex == -1 || quantityColumnIndex == -1) {
                    throw new IllegalArgumentException("Не найдены необходимые колонки в исходном файле: " + file.getName() + ". Пожалуйста, проверьте структуру файла.");
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        String articul = HelperClass.getCellStringValue(row, supplierArticulColumnIndex).trim().toLowerCase();
                        String reasonForPayment = HelperClass.getCellStringValue(row, reasonForPaymentColumnIndex).trim();
                        double saleAmount = HelperClass.getNumericCellValue(row.getCell(saleAmountColumnIndex));
                        int quantity = (int) HelperClass.getNumericCellValue(row.getCell(quantityColumnIndex));

                        if (!articul.isEmpty() && "Продажа".equalsIgnoreCase(reasonForPayment) && saleAmount > 0) {
                            for (Product product : uniqueProducts) {
                                if (product.getArticul().equals(articul)) {
                                    countMap.put(articul, countMap.getOrDefault(articul, 0) + quantity);
                                    saleSumMap.put(articul, saleSumMap.getOrDefault(articul, 0.0) + saleAmount);
                                }
                            }
                        }

                        if (!articul.isEmpty() && "Возврат".equalsIgnoreCase(reasonForPayment) && quantity == 1) {
                            for (Product product : uniqueProducts) {
                                if (product.getArticul().equals(articul)) {
                                    returnCountMap.put(articul, returnCountMap.getOrDefault(articul, 0) + quantity);
                                    returnSumMap.put(articul, returnSumMap.getOrDefault(articul, 0.0) + saleAmount);
                                }
                            }
                        }
                    }
                }
            }
        }

        String outputFilePath = getOutputReportPath();
        Workbook reportWorkbook;

        File outputFile = new File(getOutputReportPath());
        if (outputFile.exists()) {
            reportWorkbook = HelperClass.loadWorkbook(outputFile);
        } else {
            String[] headers = {"Название", "Артикул поставщика", "Кол-во", "Вайлдберриз реализовал товар (Пр)", "Возврат", "Общая сумма возврата"};
            reportWorkbook = HelperClass.createWorkbook(headers);
        }

        Sheet reportSheet = reportWorkbook.getSheetAt(0);
        if (reportSheet == null) {
            reportSheet = reportWorkbook.createSheet("Отчет");
        }

        Map<String, Object[]> reportMap = new HashMap<>();
        for (Product product : uniqueProducts) {
            int correctedCount = countMap.getOrDefault(product.getArticul(), 0) - returnCountMap.getOrDefault(product.getArticul(), 0);
            double correctedSaleSum = saleSumMap.getOrDefault(product.getArticul(), 0.0) - returnSumMap.getOrDefault(product.getArticul(), 0.0);

            Object[] rowData = {
                    product.getName(),
                    product.getArticul(),
                    correctedCount,
                    correctedSaleSum,
                    returnCountMap.getOrDefault(product.getArticul(), 0),
                    returnSumMap.getOrDefault(product.getArticul(), 0.0)
            };
            reportMap.put(product.getArticul(), rowData);
        }

        HelperClass.writeDataToSheet(reportSheet, 1, reportMap);

        HelperClass.saveWorkbook(reportWorkbook, getOutputReportPath());
        log.info("Отчет успешно обновлен с количеством и суммой продаж/возвратов: " + getOutputReportPath());
    }


    public void inputVozToNds(Set<Product> uniqueProducts) throws IOException {
        if (sourceRussianFilePath == null || sourceInternationalFilePath == null) {
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файлы перед генерацией отчета.");
        }

        Map<String, Double> vozmeshenieSumMap = new HashMap<>();
        Map<String, Double> ekvayringSumMap = new HashMap<>();
        Map<String, Double> vozWbBezNdsSumMap = new HashMap<>();
        Map<String, Double> ndsWbSumMap = new HashMap<>();
        Map<String, Double> kPerechisleniyuSumMap = new HashMap<>();

        List<File> files = List.of(new File(sourceRussianFilePath), new File(sourceInternationalFilePath));

        for (File file : files) {
            try (Workbook workbook = HelperClass.loadWorkbook(file)) {
                Sheet sheet = workbook.getSheetAt(0);

                int articulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");
                int vozmeshenieColumnIndex = HelperClass.findColumnIndex(sheet, "Возмещение за выдачу и возврат товаров на ПВЗ");
                int ekvayringColumnIndex = HelperClass.findColumnIndex(sheet, "Эквайринг/Комиссия за организацию платежей");
                int vozWbBezNdsColumnIndex = HelperClass.findColumnIndex(sheet, "Вознаграждение Вайлдберриз (ВВ), без НДС");
                int ndsWbColumnIndex = HelperClass.findColumnIndex(sheet, "НДС с Вознаграждения Вайлдберриз");
                int kPerechisleniyuColumnIndex = HelperClass.findColumnIndex(sheet, "К перечислению Продавцу за реализованный Товар");

                if (articulColumnIndex == -1 || vozmeshenieColumnIndex == -1 || ekvayringColumnIndex == -1 ||
                        vozWbBezNdsColumnIndex == -1 || ndsWbColumnIndex == -1 || kPerechisleniyuColumnIndex == -1) {
                    throw new IllegalArgumentException("Не найдены необходимые колонки в исходном файле.");
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        String articul = HelperClass.getCellStringValue(row, articulColumnIndex).trim().toLowerCase();

                        if (uniqueProducts.stream().anyMatch(product -> product.getArticul().equals(articul))) {
                            double vozmeshenieValue = HelperClass.getNumericCellValue(row.getCell(vozmeshenieColumnIndex));
                            double ekvayringValue = HelperClass.getNumericCellValue(row.getCell(ekvayringColumnIndex));
                            double vozWbBezNdsValue = HelperClass.getNumericCellValue(row.getCell(vozWbBezNdsColumnIndex));
                            double ndsWbValue = HelperClass.getNumericCellValue(row.getCell(ndsWbColumnIndex));
                            double kPerechisleniyuValue = HelperClass.getNumericCellValue(row.getCell(kPerechisleniyuColumnIndex));

                            vozmeshenieSumMap.put(articul, vozmeshenieSumMap.getOrDefault(articul, 0.0) + vozmeshenieValue);
                            ekvayringSumMap.put(articul, ekvayringSumMap.getOrDefault(articul, 0.0) + ekvayringValue);
                            vozWbBezNdsSumMap.put(articul, vozWbBezNdsSumMap.getOrDefault(articul, 0.0) + vozWbBezNdsValue);
                            ndsWbSumMap.put(articul, ndsWbSumMap.getOrDefault(articul, 0.0) + ndsWbValue);
                            kPerechisleniyuSumMap.put(articul, kPerechisleniyuSumMap.getOrDefault(articul, 0.0) + kPerechisleniyuValue);
                        }
                    }
                }
            }
        }

        String outputFilePath = getOutputReportPath();

        try (Workbook reportWorkbook = HelperClass.loadWorkbook(new File(outputFilePath))) {
            Sheet reportSheet = reportWorkbook.getSheetAt(0);

            for (Product product : uniqueProducts) {
                String articul = product.getArticul();

                Row reportRow = HelperClass.findRowByArticul(reportSheet, articul);
                if (reportRow != null) {
                    HelperClass.setNumericCellValue(reportRow, 4, vozmeshenieSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 5, ekvayringSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 6, vozWbBezNdsSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 7, ndsWbSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 8, kPerechisleniyuSumMap.getOrDefault(articul, 0.0));
                }
            }

            HelperClass.saveWorkbook(reportWorkbook, outputFilePath);
        }

        log.info("Отчет успешно обновлен с данными по возмещению и НДС: " + outputFilePath);
    }

    public void inputUslToPlat(Set<Product> uniqueProducts) throws IOException {
        if (sourceRussianFilePath == null || sourceInternationalFilePath == null) {
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файлы перед генерацией отчета.");
        }

        Map<String, Double> uslugiSumMap = new HashMap<>();
        Map<String, Double> shtrafySumMap = new HashMap<>();
        Map<String, Double> hranenieSumMap = new HashMap<>();
        Map<String, Double> uderzhaniyaSumMap = new HashMap<>();
        Map<String, Double> platnayaPriemkaSumMap = new HashMap<>();

        List<File> files = List.of(new File(sourceRussianFilePath), new File(sourceInternationalFilePath));

        for (File file : files) {
            try (Workbook workbook = HelperClass.loadWorkbook(file)) {
                Sheet sheet = workbook.getSheetAt(0);

                int articulColumnIndex = HelperClass.findColumnIndex(sheet, "Артикул поставщика");
                int uslugiColumnIndex = HelperClass.findColumnIndex(sheet, "Услуги по доставке товара покупателю");
                int shtrafyColumnIndex = HelperClass.findColumnIndex(sheet, "Общая сумма штрафов");
                int hranenieColumnIndex = HelperClass.findColumnIndex(sheet, "Хранение");
                int uderzhaniyaColumnIndex = HelperClass.findColumnIndex(sheet, "Удержания");
                int platnayaPriemkaColumnIndex = HelperClass.findColumnIndex(sheet, "Платная приемка");

                if (articulColumnIndex == -1 || uslugiColumnIndex == -1 || shtrafyColumnIndex == -1 ||
                        hranenieColumnIndex == -1 || uderzhaniyaColumnIndex == -1 || platnayaPriemkaColumnIndex == -1) {
                    throw new IllegalArgumentException("Не найдены необходимые колонки в исходном файле.");
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        String articul = HelperClass.getCellStringValue(row, articulColumnIndex).trim().toLowerCase();

                        if (uniqueProducts.stream().anyMatch(product -> product.getArticul().equals(articul))) {
                            double uslugiValue = HelperClass.getNumericCellValue(row.getCell(uslugiColumnIndex));
                            double shtrafyValue = HelperClass.getNumericCellValue(row.getCell(shtrafyColumnIndex));
                            double hranenieValue = HelperClass.getNumericCellValue(row.getCell(hranenieColumnIndex));
                            double uderzhaniyaValue = HelperClass.getNumericCellValue(row.getCell(uderzhaniyaColumnIndex));
                            double platnayaPriemkaValue = HelperClass.getNumericCellValue(row.getCell(platnayaPriemkaColumnIndex));

                            uslugiSumMap.put(articul, uslugiSumMap.getOrDefault(articul, 0.0) + uslugiValue);
                            shtrafySumMap.put(articul, shtrafySumMap.getOrDefault(articul, 0.0) + shtrafyValue);
                            hranenieSumMap.put(articul, hranenieSumMap.getOrDefault(articul, 0.0) + hranenieValue);
                            uderzhaniyaSumMap.put(articul, uderzhaniyaSumMap.getOrDefault(articul, 0.0) + uderzhaniyaValue);
                            platnayaPriemkaSumMap.put(articul, platnayaPriemkaSumMap.getOrDefault(articul, 0.0) + platnayaPriemkaValue);
                        }
                    }
                }
            }
        }

        String outputFilePath = getOutputReportPath();

        try (Workbook reportWorkbook = HelperClass.loadWorkbook(new File(outputFilePath))) {
            Sheet reportSheet = reportWorkbook.getSheetAt(0);

            for (Product product : uniqueProducts) {
                String articul = product.getArticul();

                Row reportRow = HelperClass.findRowByArticul(reportSheet, articul);
                if (reportRow != null) {
                    HelperClass.setNumericCellValue(reportRow, 9, uslugiSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 10, shtrafySumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 11, hranenieSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 12, uderzhaniyaSumMap.getOrDefault(articul, 0.0));
                    HelperClass.setNumericCellValue(reportRow, 13, platnayaPriemkaSumMap.getOrDefault(articul, 0.0));
                }
            }

            HelperClass.saveWorkbook(reportWorkbook, outputFilePath);
        }

        log.info("Отчет успешно обновлен с данными по услугам, штрафам, хранению, удержаниям и платной приемке: " + outputFilePath);
    }

    public void inputOplToPrib(Set<Product> uniqueProducts) throws IOException {
        if (sourceRussianFilePath == null || sourceInternationalFilePath == null) {
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файлы перед генерацией отчета.");
        }

        String outputFilePath = getOutputReportPath();

        try (Workbook reportWorkbook = HelperClass.loadWorkbook(new File(getOutputReportPath()))) {
            Sheet reportSheet = reportWorkbook.getSheetAt(0);

            LocalDate minDate = null;
            LocalDate maxDate = null;

            try (Workbook russianWorkbook = HelperClass.loadWorkbook(new File(sourceRussianFilePath));
                 Workbook internationalWorkbook = HelperClass.loadWorkbook(new File(sourceInternationalFilePath))) {

                minDate = findDateRange(russianWorkbook,  minDate, true);
                maxDate = findDateRange(russianWorkbook,  maxDate, false);

                minDate = findDateRange(internationalWorkbook,  minDate, true);
                maxDate = findDateRange(internationalWorkbook,  maxDate, false);
            }

            for (Product product : uniqueProducts) {
                String articul = product.getArticul().toLowerCase();

                Row reportRow = HelperClass.findRowByArticul(reportSheet, articul);
                if (reportRow != null) {
                    double kPerechisleniyu = HelperClass.getNumericCellValue(reportRow.getCell(8));
                    double shtrafy = HelperClass.getNumericCellValue(reportRow.getCell(10));
                    double hranenie = HelperClass.getNumericCellValue(reportRow.getCell(11));
                    double uderzhaniya = HelperClass.getNumericCellValue(reportRow.getCell(12));
                    double platnayaPriemka = HelperClass.getNumericCellValue(reportRow.getCell(13));
                    double uslugiPoDostavke = HelperClass.getNumericCellValue(reportRow.getCell(9));
                    double kolvo = HelperClass.getNumericCellValue(reportRow.getCell(2));

                    double kOplate = kPerechisleniyu - (shtrafy + hranenie + uderzhaniya + platnayaPriemka + uslugiPoDostavke);
                    HelperClass.setNumericCellValue(reportRow, 14, kOplate);

                    double sebestoimostNaEdinicu = productRepository.findCostByArticulAndUser(articul, product.getUser());

                    double sebestoimost = sebestoimostNaEdinicu * kolvo;
                    HelperClass.setNumericCellValue(reportRow, 15, sebestoimost);

                    double pribylDoNaloga = kOplate - sebestoimost;
                    HelperClass.setNumericCellValue(reportRow, 16, pribylDoNaloga);
                }
            }

            if (minDate != null && maxDate != null) {
                log.info("Добавление строки с периодом продаж: {} - {}", minDate, maxDate);
                int lastRowNum = reportSheet.getLastRowNum();
                Row periodRow = reportSheet.createRow(lastRowNum + 1);
                Cell periodCell = periodRow.createCell(0);
                periodCell.setCellValue("Период продаж: " + minDate.toString() + " - " + maxDate.toString());
            } else {
                log.warn("Период продаж не определен. Даты: minDate={}, maxDate={}", minDate, maxDate);
            }

            HelperClass.saveWorkbook(reportWorkbook, outputFilePath);
        }

        log.info("Колонки 'К оплате', 'Себестоимость', 'Прибыль до налогообложения' успешно рассчитаны, а период продаж добавлен: " + outputFilePath);
    }

    public void addSummaryRowToFileWithExternalData() throws IOException {
        String outputFilePath = getOutputReportPath();

        double extraHranenieSum = 0.0;
        double extraUderzhaniyaSum = 0.0;
        double extraPlatnayaPriemkaSum = 0.0;

        List<File> sourceFiles = List.of(getRussianFile(), getInternationalFile());

        try {
            for (File sourceFile : sourceFiles) {
                try (Workbook sourceWorkbook = HelperClass.loadWorkbook(sourceFile)) {
                    Sheet sourceSheet = sourceWorkbook.getSheetAt(0);

                    int hranenieColumnIndex = HelperClass.findColumnIndex(sourceSheet, "Хранение");
                    int uderzhaniyaColumnIndex = HelperClass.findColumnIndex(sourceSheet, "Удержания");
                    int platnayaPriemkaColumnIndex = HelperClass.findColumnIndex(sourceSheet, "Платная приемка");

                    if (hranenieColumnIndex == -1 || uderzhaniyaColumnIndex == -1 || platnayaPriemkaColumnIndex == -1) {
                        throw new IllegalArgumentException("Не найдены необходимые колонки в исходных файлах.");
                    }

                    for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
                        Row row = sourceSheet.getRow(i);
                        if (row != null) {
                            extraHranenieSum += HelperClass.getNumericCellValue(row.getCell(hranenieColumnIndex));
                            extraUderzhaniyaSum += HelperClass.getNumericCellValue(row.getCell(uderzhaniyaColumnIndex));
                            extraPlatnayaPriemkaSum += HelperClass.getNumericCellValue(row.getCell(platnayaPriemkaColumnIndex));
                        }
                    }
                }
            }

            try (Workbook workbook = HelperClass.loadWorkbook(new File(outputFilePath))) {
                Sheet sheet = workbook.getSheetAt(0);

                int lastRowNum = sheet.getLastRowNum();
                Row summaryRow = sheet.createRow(lastRowNum + 1);

                Cell summaryLabelCell = summaryRow.createCell(0);
                summaryLabelCell.setCellValue("Итог:");

                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                style.setFont(font);
                summaryLabelCell.setCellStyle(style);

                for (int colIndex = 1; colIndex < sheet.getRow(0).getLastCellNum(); colIndex++) {
                    double sum = 0.0;

                    for (int rowIndex = 1; rowIndex <= lastRowNum; rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        if (row != null) {
                            Cell cell = row.getCell(colIndex);
                            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                                sum += cell.getNumericCellValue();
                            }
                        }
                    }

                    String columnName = sheet.getRow(0).getCell(colIndex).getStringCellValue().trim();
                    if ("Хранение".equalsIgnoreCase(columnName)) {
                        sum += extraHranenieSum;
                    } else if ("Удержания".equalsIgnoreCase(columnName)) {
                        sum += extraUderzhaniyaSum;
                    } else if ("Платная приемка".equalsIgnoreCase(columnName)) {
                        sum += extraPlatnayaPriemkaSum;
                    }

                    Cell summaryCell = summaryRow.createCell(colIndex);
                    summaryCell.setCellValue(sum);
                    summaryCell.setCellStyle(style);
                }

                HelperClass.saveWorkbook(workbook, outputFilePath);
                log.info("Строка итогов с внешними данными успешно добавлена в файл: " + outputFilePath);
            }
        } finally {
            // Удаление временных файлов
            for (File file : sourceFiles) {
                if (file.exists() && !file.delete()) {
                    log.warn("Не удалось удалить временный файл: " + file.getAbsolutePath());
                } else {
                    log.info("Удалён временный файл: " + file.getAbsolutePath());
                }
            }
        }
    }




    @Transactional
    public void updateCosts(Map<Long, Double> costsMap) {
        for (Map.Entry<Long, Double> entry : costsMap.entrySet()) {
            Long productId = entry.getKey();
            Double newCost = entry.getValue();

            Optional<Product> optionalProduct = productRepository.findById(productId);
            if (optionalProduct.isPresent()) {
                Product product = optionalProduct.get();
                log.info("Перед обновлением продукта ID {}: {}", productId, product);

                product.setCost(newCost);
                product.setUpdateCost(LocalDateTime.now());

                productRepository.save(product);
                log.info("Обновили себестоимость для продукта ID : {}: {}", productId, newCost);
            } else {
                log.warn("Продукт не найден с ID: {}", productId);
            }
        }
    }



    public File getInternationalFile() {
        if (sourceInternationalFilePath == null) {
            log.error("Не найден исходный файл");
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файл перед генерацией отчета.");
        }
        return new File(sourceInternationalFilePath);
    }

    public File getRussianFile() {
        if (sourceRussianFilePath == null) {
            log.error("Не найден исходный файл");
            throw new IllegalStateException("Исходный файл не установлен. Загрузите файл перед генерацией отчета.");
        }
        return new File(sourceRussianFilePath);
    }

    private LocalDate findDateRange(Workbook workbook, LocalDate currentDate, boolean findMin) {
        final String targetColumnName = "Дата продажи";
        Sheet sheet = workbook.getSheetAt(0);
        int dateColumnIndex = HelperClass.findColumnIndex(sheet, targetColumnName);

        if (dateColumnIndex == -1) {
            log.error("Колонка '{}' не найдена в файле.", targetColumnName);
            throw new IllegalArgumentException("Не найдена колонка '" + targetColumnName + "' в исходном файле.");
        }

        log.info("Колонка '{}' найдена. Индекс: {}", targetColumnName, dateColumnIndex);

        LocalDate minDate = null;
        LocalDate maxDate = null;

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(dateColumnIndex);
                LocalDate date = HelperClass.getDateCellValue(cell);
                log.debug("Строка {}: Тип ячейки: {}, Значение: {}, Преобразованная дата: {}",
                        i, (cell != null ? cell.getCellType() : "null"),
                        (cell != null ? cell.toString() : "null"), date);

                if (date == null) {
                    log.warn("Ячейка в строке {} и колонке {} не содержит валидной даты.", i, dateColumnIndex);
                    continue;
                }

                // Фильтрация некорректных или старых дат
                if (date.isBefore(LocalDate.of(2000, 1, 1)) || date.isAfter(LocalDate.now().plusYears(1))) {
                    log.warn("Игнорируем дату {}, так как она вне допустимого диапазона.", date);
                    continue;
                }

                if (minDate == null || date.isBefore(minDate)) {
                    minDate = date;
                }
                if (maxDate == null || date.isAfter(maxDate)) {
                    maxDate = date;
                }
            }
        }

        log.info("Минимальная дата: {}, Максимальная дата: {}", minDate, maxDate);
        return findMin ? minDate : maxDate;
    }


    public void generateMonthlyReport(MultipartFile[] weeklyReports, String month, String outputDirectory) throws IOException {
        User currentUser = customUserDetailService.getCurrentUser();

        Report report = new Report();
        String outputFilePath = outputDirectory + "/" + month + ".xlsx";
        report.setFilePath(outputFilePath);
        report.setFilename(month);
        report.setUser(currentUser);
        report.setLocalDateTime(LocalDateTime.now());
        reportRepository.save(report);

        File outputDir = new File(outputDirectory);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
            log.info("Папка для сохранения отчёта не существует. Создаём: " + outputDirectory);
        }

        Workbook monthlyWorkbook = new XSSFWorkbook();
        Sheet sheet = monthlyWorkbook.createSheet("Monthly Report");

        // Создаём стили для заголовков и итоговой строки
        CellStyle headerStyle = monthlyWorkbook.createCellStyle();
        Font headerFont = monthlyWorkbook.createFont();
        headerFont.setBold(true); // Жирный шрифт
        headerFont.setColor(IndexedColors.BLACK.getIndex()); // Чёрный цвет текста
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIME.getIndex()); // Салатовый фон
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // Устанавливаем сплошной фон

        CellStyle totalRowStyle = monthlyWorkbook.createCellStyle();
        Font totalRowFont = monthlyWorkbook.createFont();
        totalRowFont.setBold(true); // Жирный шрифт
        totalRowStyle.setFont(totalRowFont);

        // Заголовки отчёта
        String[] headers = {"Дата", "Продажи", "Разница после комиссий",
                "К перечислению продавцу за реализованный товар", "Логистика", "Штрафы", "Стоимость платной приемки",
                "Стоимость хранения", "Прочие удержания/Выплаты(продвижение ВБ и т.д.)",
                "Логистика, хранение, прочие удержания", "К оплате",
                "Себестоимость", "Прибыль до налогообложения", "Налог УСН", "Чистая прибыль после УСН"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle); // Применяем стиль для заголовков
        }

        int totalColumns = headers.length;
        for (int i = 0; i < totalColumns; i++) {
            sheet.setColumnWidth(i, 8000);
        }

        int rowHeightInPoints = 30;
        headerRow.setHeightInPoints(rowHeightInPoints);

        int rowNum = 1;
        double totalSales = 0;
        double totalDifference = 0;
        double totalTransfer = 0;
        double totalLogistic = 0;
        double totalShtrafov = 0;
        double totalPlatnayaPriemka = 0;
        double totalXranenie = 0;
        double totalYderzhania = 0;
        double totalLogisticXranenieYderzhania = 0;
        double totalKOplate = 0;
        double totalSebestoimost = 0;
        double totalPribilDoNaloga = 0;
        double totalNalogUsn = 0;
        double totalChistayaPribilPosleUsn = 0;

        for (MultipartFile file : weeklyReports) {
            Workbook weeklyWorkbook = new XSSFWorkbook(new ByteArrayInputStream(file.getBytes()));
            Sheet weeklySheet = weeklyWorkbook.getSheetAt(0);

            int lastRowNum = weeklySheet.getLastRowNum();
            int predLastRowNum = weeklySheet.getLastRowNum() - 1;
            Row lastRow = weeklySheet.getRow(lastRowNum);
            Row predLastRow = weeklySheet.getRow(predLastRowNum);

            if (lastRow != null) {
                Row monthlyRow = sheet.createRow(rowNum++);
                monthlyRow.setHeightInPoints(rowHeightInPoints);
                String date = HelperClass.getCellStringValue(predLastRow, findColumnIndex(weeklySheet, "Название"));
                monthlyRow.createCell(0).setCellValue(date);
                double sales = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Вайлдберриз реализовал Товар (Пр)")), 0.0);
                double transfer = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "К перечислению Продавцу за реализованный Товар")), 0.0);
                double logistic = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Услуги по доставке товара покупателю")), 0.0);
                double sumShtrafov = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Общая сумма штрафов")), 0.0);
                double difference = transfer - sales;
                double platnayaPriemka = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Платная приемка")), 0.0);
                double xranenie = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Хранение")), 0.0);
                double yderzhania = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "Удержания")), 0.0);
                double logisticXranenieYderzhania = logistic + sumShtrafov + platnayaPriemka + xranenie + yderzhania;
                double kOplate = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "к оплате")), 0.0);
                double sebestoimost = HelperClass.getNumericCellValueOrDefault(lastRow.getCell(findColumnIndex(weeklySheet, "к оплате")), 0.0);
                double pribilDoNaloga = kOplate - sebestoimost;
                double nalogUsn = sales * 0.015;
                double chistayaPribilPosleUsn = pribilDoNaloga - nalogUsn;

                // Добавляем значения в строку
                monthlyRow.createCell(1).setCellValue(sales);
                monthlyRow.createCell(2).setCellValue(difference);
                monthlyRow.createCell(3).setCellValue(transfer);
                monthlyRow.createCell(4).setCellValue(logistic);
                monthlyRow.createCell(5).setCellValue(sumShtrafov);
                monthlyRow.createCell(6).setCellValue(platnayaPriemka);
                monthlyRow.createCell(7).setCellValue(xranenie);
                monthlyRow.createCell(8).setCellValue(yderzhania);
                monthlyRow.createCell(9).setCellValue(logisticXranenieYderzhania);
                monthlyRow.createCell(10).setCellValue(kOplate);
                monthlyRow.createCell(11).setCellValue(sebestoimost);
                monthlyRow.createCell(12).setCellValue(pribilDoNaloga);
                monthlyRow.createCell(13).setCellValue(nalogUsn);
                monthlyRow.createCell(14).setCellValue(chistayaPribilPosleUsn);

                // Суммируем значения для итогов
                totalSales += sales;
                totalDifference += difference;
                totalTransfer += transfer;
                totalLogistic += logistic;
                totalShtrafov += sumShtrafov;
                totalPlatnayaPriemka += platnayaPriemka;
                totalXranenie += xranenie;
                totalYderzhania += yderzhania;
                totalLogisticXranenieYderzhania += logisticXranenieYderzhania;
                totalKOplate += kOplate;
                totalSebestoimost += sebestoimost;
                totalPribilDoNaloga += pribilDoNaloga;
                totalNalogUsn += nalogUsn;
                totalChistayaPribilPosleUsn += chistayaPribilPosleUsn;
            }
        }

        // Добавляем итоговую строку с жирным шрифтом
        Row totalRow = sheet.createRow(rowNum++);
        totalRow.createCell(0).setCellValue("Итог");
        totalRow.createCell(1).setCellValue(totalSales);
        totalRow.createCell(2).setCellValue(totalDifference);
        totalRow.createCell(3).setCellValue(totalTransfer);
        totalRow.createCell(4).setCellValue(totalLogistic);
        totalRow.createCell(5).setCellValue(totalShtrafov);
        totalRow.createCell(6).setCellValue(totalPlatnayaPriemka);
        totalRow.createCell(7).setCellValue(totalXranenie);
        totalRow.createCell(8).setCellValue(totalYderzhania);
        totalRow.createCell(9).setCellValue(totalLogisticXranenieYderzhania);
        totalRow.createCell(10).setCellValue(totalKOplate);
        totalRow.createCell(11).setCellValue(totalSebestoimost);
        totalRow.createCell(12).setCellValue(totalPribilDoNaloga);
        totalRow.createCell(13).setCellValue(totalNalogUsn);
        totalRow.createCell(14).setCellValue(totalChistayaPribilPosleUsn);

        // Применяем жирный шрифт к итоговой строке
        for (int i = 0; i < totalColumns; i++) {
            totalRow.getCell(i).setCellStyle(totalRowStyle);
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
            monthlyWorkbook.write(fileOut);
        }

        log.info("Месячный отчёт успешно создан: " + outputFilePath);
    }





    public static int findColumnIndex(Sheet sheet, String columnName) { Row headerRow = sheet.getRow(0); for (Cell cell : headerRow) { if (cell.getStringCellValue().trim().equalsIgnoreCase(columnName)) { return cell.getColumnIndex(); } } log.error("Не удалось найти колонку '{}'", columnName); return -1; }


}






