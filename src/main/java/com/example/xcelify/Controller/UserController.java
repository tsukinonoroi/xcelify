package com.example.xcelify.Controller;


import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
@RequiredArgsConstructor
public class UserController {
    @GetMapping("/login")
    public String login() {
        return "login";
    }

    @GetMapping("/")
    public String main() {
        return "main";
    }

}
