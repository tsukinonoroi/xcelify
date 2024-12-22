package com.example.xcelify.Repository;

import com.example.xcelify.Model.Report;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;


@Repository
public interface ReportRepository extends JpaRepository<Report, Long> {

    List<Report> findAllByUser_Id(Long id);

    Optional<Report> findByFilename(String filename);
}
