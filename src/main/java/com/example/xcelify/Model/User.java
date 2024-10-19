package com.example.xcelify.Model;

import jakarta.persistence.*;
import lombok.Data;
import lombok.Generated;
import org.springframework.context.annotation.Primary;

@Entity
@Table(name = "user")
@Data
public class User {
    @Id
    @Column(name = "id")
    private int id;

    @Column(name = "login")
    private String login;

    @Column(name = "password")
    private String password;
}
