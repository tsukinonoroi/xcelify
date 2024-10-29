package com.example.xcelify.Model;

import jakarta.persistence.*;
import lombok.Data;

import java.time.LocalDate;
import java.time.LocalDateTime;

@Entity
@Data
@Table(name = "reports")
public class Report {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @Column(name = "filename")
    private String filename;

    @Column(name = "localdatetime")
    private LocalDateTime localDateTime;

    @Column(name = "filePath")
    private String filePath;

}
