package com.example.xcelify.Model;

import jakarta.persistence.*;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Data
@Entity
@Table(name = "ProductHistory")
public class ProductHistory {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @ManyToOne
    @JoinColumn(name = "product_id", nullable = false)
    private Product product;

    @Column(name = "cost")
    private Double cost;

    @Column(name = "update_time")
    private LocalDateTime updateTime;

    public ProductHistory(Product product, Double cost, LocalDateTime updateTime) {
        this.product = product;
        this.cost = cost;
        this.updateTime = updateTime;
    }

    public ProductHistory() {
    }
}