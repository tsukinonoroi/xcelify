package com.example.xcelify.Controller;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Service.ReportService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;


import java.io.IOException;
import java.util.*;

@Slf4j
@Controller
@RequiredArgsConstructor
public class UploadController {

    private final ReportService reportService;
    private final ProductRepository productRepository;

    @PostMapping("/upload")
    public String uploadFile(@RequestParam("file") MultipartFile file, Model model) throws IOException {
        Set<Product> productsWithCosts = reportService.parseUniqueProducts(file);
        List<Product> products = new ArrayList<>(productsWithCosts);
        model.addAttribute("products", products);
        return "enter_costs";
    }

    @GetMapping("/enter_costs")
    public String enterCosts(Model model) {
        List<Product> products = productRepository.findAll(); 
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

        return "redirect:/enter_costs";
    }
}
