package com.example.xcelify.Controller;

import com.example.xcelify.Service.HelperClass;
import com.example.xcelify.Service.ReportService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

@Controller
@RequiredArgsConstructor
public class MonthlyController {

    private final ReportService reportService;
    private final String filePath = "C:/users/edemw/Desktop/filter";

    @PostMapping("/monthly-report")
    public String generateMonthlyReport(
            @RequestParam("month") String month,
            @RequestParam("fileWeekly1") MultipartFile fileWeekly1,
            @RequestParam("fileWeekly2") MultipartFile fileWeekly2,
            @RequestParam("fileWeekly3") MultipartFile fileWeekly3,
            @RequestParam("fileWeekly4") MultipartFile fileWeekly4) {
        try {
            MultipartFile[] weeklyArray = {fileWeekly1, fileWeekly2, fileWeekly3, fileWeekly4};

            reportService.generateMonthlyReport(weeklyArray, month, filePath);

            return "redirect:/reports";
        } catch (Exception e) {
            return "redirect:/error?message=" + e.getMessage();
        }
    }


}
