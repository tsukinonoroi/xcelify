package com.example.xcelify.Model;

import jakarta.annotation.Nullable;
import jakarta.persistence.*;
import lombok.Data;

import java.time.LocalDateTime;

@Data
@Entity
@Table(name = "Product")
public class Product {
    @Id
    @GeneratedValue (strategy = GenerationType.AUTO)
    private Long id;
    @Column(name = "name")
    private String name;
    @Column(name = "articul")
    private String articul;
    @Column(name = "cost")
    @Nullable
    private Double cost = null;
    @Column(name = "updateCost")
    private LocalDateTime localDateTime;
    
}
