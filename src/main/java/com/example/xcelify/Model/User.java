package com.example.xcelify.Model;

import jakarta.persistence.*;
import lombok.Data;
import lombok.ToString;

import java.util.List;

@Entity
@Table(name = "user")
@Data
@ToString(exclude = "products")
public class User {
    @Id
    @Column(name = "id")
    private Long id;

    @Column(name = "login")
    private String login;

    @Column(name = "password")
    private String password;

    @OneToMany(mappedBy = "user")
    private List<Product> products;

    @OneToMany(mappedBy = "user")
    private List<Report> reports;
}
