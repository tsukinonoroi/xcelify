package com.example.xcelify.Model;


import jakarta.annotation.Nullable;
import jakarta.persistence.*;
import lombok.Data;
import lombok.ToString;

import java.time.LocalDateTime;
import java.util.List;

@Data
@Entity
@Table(name = "Product")
@ToString(exclude = "user")
public class Product {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @Column(name = "name")
    private String name;

    @Column(name = "articul")
    private String articul;

    @Column(name = "cost")
    @Nullable
    private Double cost = null;

    @Column(name = "updateCost")
    private LocalDateTime updateCost;

    @OneToMany(mappedBy = "product", cascade = CascadeType.ALL)
    private List<ProductHistory> history;

    @ManyToOne
    @JoinColumn(name = "user_id", nullable = false)
    private User user;

}