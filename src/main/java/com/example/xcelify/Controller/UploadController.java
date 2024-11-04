package com.example.xcelify.Controller;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Model.Report;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Service.ReportService;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;


import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

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
    public String enterCosts() {
        return "enter_costs";
    }
}
