package com.example.xcelify.Controller;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Service.ReportService;

import jakarta.servlet.http.HttpServletResponse;
import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.server.ResponseStatusException;


import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.*;

@Slf4j
@Controller
@RequiredArgsConstructor
@Data
public class UploadController {

    private final ReportService reportService;
    private final ProductRepository productRepository;


    @PostMapping("/upload")
    public String uploadFile(@RequestParam("file") MultipartFile file, Model model) throws IOException {

        String uploadDir = "C:\\Users\\edemw\\Desktop\\notfilter\\" + file.getOriginalFilename();
        File dest = new File(uploadDir);

        File parentDir = dest.getParentFile();
        if (!parentDir.exists()) {
            parentDir.mkdirs();
        }

        file.transferTo(dest);

        reportService.setSourceFilePath(uploadDir);

        Set<Product> productsWithCosts = reportService.parseUniqueProducts(dest);

        List<Product> products = new ArrayList<>(productsWithCosts);
        model.addAttribute("products", products);

        return "enter_costs";
    }

    @PostMapping("/updateCosts")
    public String updateCosts(@RequestParam Map<String, String> allParams) {
        Map<Long, Double> costsMap = new HashMap<>();

        log.debug("Received parameters: {}", allParams);

        allParams.forEach((key, value) -> {
            if (key.startsWith("costs[")) {
                String idString = key.substring(6, key.length() - 1);
                try {
                    Long productId = Long.valueOf(idString.trim().replaceAll("[^\\d]", "")); // Удаляем пробелы из ID
                    Double cost = Double.valueOf(value.replaceAll("[^\\d.]", "").replace(",", "."));

                    log.debug("Adding product ID: {}, cost: {}", productId, cost);

                    costsMap.put(productId, cost);
                } catch (NumberFormatException e) {
                    log.warn("Неверный формат ID или стоимости: ID={}, стоимость={}", idString, value);
                }
            }
        });
        reportService.updateCosts(costsMap);

        return "redirect:/updateCosts";
    }

    @PostMapping("/generateReport")
    public String generateReport(@RequestParam("reportName") String reportName) throws IOException {

        reportService.setReportName(reportName);
        reportService.generateNewReport(reportName);

        File sourceFile = reportService.getSourceFile();
        Set<Product> uniqueProducts = reportService.parseUniqueProducts(sourceFile);

        reportService.inputNameAndArticul(uniqueProducts);
        reportService.inputCountAndSale(uniqueProducts);
        return "generateReport";
    }
    @GetMapping("/generateReport")
    public String generareReportGet() {
        return "generateReport";
    }

    @GetMapping("/updateCosts")
    public String getUpdCost(Model model){
        return "enter_costs";
    }


}
