package com.example.xcelify.Controller;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Service.ReportService;

import jakarta.servlet.http.HttpServletResponse;
import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.*;

@Slf4j
@Controller
@Data
@RequiredArgsConstructor
public class UploadController {

    private final ReportService reportService;
    private final ProductRepository productRepository;


    @PostMapping("/upload")
    public String uploadFile(@RequestParam("fileRussia") MultipartFile fileRussia,
                             @RequestParam("fileInternational") MultipartFile fileInternational,
                             Model model) throws IOException {

        // Указание директории для сохранения
        String uploadDir = "/root/xcelify/notfilter";
        File uploadDirectory = new File(uploadDir);

        // Создание родительской папки, если она не существует
        if (!uploadDirectory.exists()) {
            log.warn("Родительской папки не существует, создаем папку");
            uploadDirectory.mkdirs();
        }

        // Преобразование MultipartFile в обычный файл на сервере
        File russianFile = new File(uploadDirectory, fileRussia.getOriginalFilename());
        File internationalFile = new File(uploadDirectory, fileInternational.getOriginalFilename());

        log.info("Сохраняем файл России по пути: " + russianFile.getAbsolutePath());
        log.info("Сохраняем файл International по пути: " + internationalFile.getAbsolutePath());

        // Перенос содержимого из MultipartFile в File
        fileRussia.transferTo(russianFile);
        fileInternational.transferTo(internationalFile);

        // Установка путей в сервис
        reportService.setSourceRussianFilePath(russianFile.getAbsolutePath());
        reportService.setSourceInternationalFilePath(internationalFile.getAbsolutePath());

        // Парсинг продуктов с учетом файлов
        Set<Product> productsWithCosts = reportService.parseUniqueProducts(russianFile, internationalFile);

        // Преобразуем в список и передаем на фронтенд
        List<Product> products = new ArrayList<>(productsWithCosts);
        model.addAttribute("products", products);

        return "enter_costs";
    }


    @PostMapping("/updateCosts")
    public String updateCosts(@RequestParam Map<String, String> allParams) {
        Map<Long, Double> costsMap = new HashMap<>();

        log.debug("Проверенные параметры: {} ", allParams);

        allParams.forEach((key, value) -> {
            if (key.startsWith("costs[")) {
                String idString = key.substring(6, key.length() - 1);
                try {
                    Long productId = Long.valueOf(idString.trim().replaceAll("[^\\d]", "")); // Удаляем пробелы из ID
                    Double cost = Double.valueOf(value.replaceAll("[^\\d.]", "").replace(",", "."));

                    log.info("Добавили продукт ID: {}, себестоимость: {}", productId, cost);

                    costsMap.put(productId, cost);
                } catch (NumberFormatException e) {
                    log.error("Неверный формат ID или стоимости: ID={}, стоимость={}", idString, value);
                }
            }
        });
        reportService.updateCosts(costsMap);

        return "redirect:/updateCosts";
    }

    @PostMapping("/products/update")
    public String updateCostsInProducts(@RequestParam Map<String, String> allParams) {
        Map<Long, Double> costsMap = new HashMap<>();

        log.debug("Проверенные параметры: {} ", allParams);

        allParams.forEach((key, value) -> {
            if (key.startsWith("costs[")) {
                String idString = key.substring(6, key.length() - 1);
                try {
                    Long productId = Long.valueOf(idString.trim().replaceAll("[^\\d]", "")); // Удаляем пробелы из ID
                    Double cost = Double.valueOf(value.replaceAll("[^\\d.]", "").replace(",", "."));

                    log.info("Добавили продукт ID: {}, себестоимость: {}", productId, cost);

                    costsMap.put(productId, cost);
                } catch (NumberFormatException e) {
                    log.error("Неверный формат ID или стоимости: ID={}, стоимость={}", idString, value);
                }
            }
        });
        reportService.updateCosts(costsMap);

        return "redirect:/products";
    }

    @PostMapping("/generateReport")
    public String generateReport(@RequestParam("reportName") String reportName) throws IOException {

        reportService.setReportName(reportName);
        reportService.generateNewReport(reportName);

        File internationalFile = reportService.getInternationalFile();
        File russianFile = reportService.getRussianFile();
        Set<Product> uniqueProducts = reportService.parseUniqueProducts(internationalFile, russianFile);

        reportService.inputNameAndArticul(uniqueProducts);
        reportService.inputCountAndSale(uniqueProducts);
        reportService.inputVozToNds(uniqueProducts);
        reportService.inputUslToPlat(uniqueProducts);
        reportService.inputOplToPrib(uniqueProducts);
        reportService.addSummaryRowToFileWithExternalData();

        return "redirect:/reports";
    }




    @GetMapping("/generateReport")
    public String generareReportGet() {
        return "generateReport";
    }

    @GetMapping("/updateCosts")
    public String getUpdCost(Model model) {
        try {
            File internationalFile = reportService.getInternationalFile();
            File russianFile = reportService.getRussianFile();

            Set<Product> uniqueProducts = reportService.parseUniqueProducts(russianFile, internationalFile);
            List<Product> uniqueProductsList = new ArrayList<>(uniqueProducts);
            Collections.sort(uniqueProductsList, Comparator.comparing(Product::getName)); // Пример сортировки по имени
            model.addAttribute("products", uniqueProductsList);

        } catch (IOException e) {
            log.error("Ошибка при загрузке исходного файла: ", e);
            model.addAttribute("error", "Не удалось загрузить данные из файла.");
        }
        return "enter_costs";
    }


}
