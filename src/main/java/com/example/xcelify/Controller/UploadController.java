package com.example.xcelify.Controller;

import com.example.xcelify.Model.Report;
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

    @Autowired
    private final ReportService reportService;

    @PostMapping("/upload")
    public String handleFileUpload(@RequestParam("report") MultipartFile file, Model model) throws IOException {
        Set<String> uniqueProductSet = reportService.parseUniqueProducts(file);

        List<String> uniqueProducts = new ArrayList<>(uniqueProductSet);
        model.addAttribute("uniqueProducts", uniqueProducts);

        return "enter_costs";
    }

    @GetMapping("/enter_costs")
    public String showInputCostsForm() {
        return "enter_costs";
    }
}
