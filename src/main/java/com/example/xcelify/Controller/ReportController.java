package com.example.xcelify.Controller;

import com.example.xcelify.Model.Report;
import com.example.xcelify.Model.User;
import com.example.xcelify.Repository.ReportRepository;
import com.example.xcelify.Service.CustomUserDetailService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.File;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.List;

@Slf4j
@Controller
@RequiredArgsConstructor
public class ReportController {
    private final CustomUserDetailService customUserDetailService;
    private final ReportRepository reportRepository;


    @GetMapping("/reports")
    public String getReports(Model model) {
        User currentUser = customUserDetailService.getCurrentUser();


        List<Report> userReports = reportRepository.findAllByUser_Id(currentUser.getId());
        userReports.sort(Comparator.comparing(Report::getLocalDateTime).reversed());
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm");
        for (Report report : userReports) {
            report.setFormattedDate(report.getLocalDateTime().format(formatter));
        }

        model.addAttribute("reports", userReports);

        return "reports";
    }

    @PostMapping("/report/delete/{filename}")
    public String deleteReportByFilename(@PathVariable String filename, RedirectAttributes redirectAttributes) {
        log.info("Удаление отчета с именем: " + filename);
        try {
            Report report = reportRepository.findByFilename(filename)
                    .orElseThrow(() -> new IllegalArgumentException("Отчёт с именем " + filename + " не найден"));

            log.info("Отчёт найден: " + report.getFilename());

            File file = new File(report.getFilePath());

            if (file.exists()) {
                boolean deleted = file.delete();
                if (deleted) {
                    log.info("Файл удалён: " + report.getFilePath());
                } else {
                    log.warn("Не удалось удалить файл: " + report.getFilePath());
                    redirectAttributes.addFlashAttribute("error", "Не удалось удалить файл.");
                    return "redirect:/reports";
                }
            } else {
                log.warn("Файл не найден: " + report.getFilePath());
                redirectAttributes.addFlashAttribute("error", "Файл не найден.");
                return "redirect:/reports";
            }

            reportRepository.delete(report);
            log.info("Запись удалена из БД: " + report.getFilename());

            redirectAttributes.addFlashAttribute("success", "Отчёт успешно удалён.");
        } catch (Exception e) {
            log.error("Ошибка при удалении отчёта: ", e);
            redirectAttributes.addFlashAttribute("error", "Не удалось удалить отчёт.");
        }
        return "redirect:/reports";
    }


    @GetMapping("/report/download/{id}")
    public ResponseEntity<Resource> downloadReport(@PathVariable String id) {
        try {
            Long cleanId = Long.parseLong(id.replace("\u00A0", "").replaceAll("\\s+", ""));

            Report report = reportRepository.findById(cleanId)
                    .orElseThrow(() -> new IllegalArgumentException("Отчёт с id " + cleanId + " не найден"));

            File file = new File(report.getFilePath());
            if (!file.exists()) {
                throw new IllegalArgumentException("Файл с путём " + report.getFilePath() + " не найден");
            }

            String encodedFilename = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8.toString())
                    .replace("+", "%20");

            Resource resource = new FileSystemResource(file);
            HttpHeaders headers = new HttpHeaders();

            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + encodedFilename);

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(file.length())
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(resource);
        } catch (NumberFormatException e) {
            log.error("Неверный формат ID: ", e);
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        } catch (Exception e) {
            log.error("Ошибка при скачивании отчёта: ", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }



}
