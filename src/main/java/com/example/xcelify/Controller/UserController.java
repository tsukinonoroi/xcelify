package com.example.xcelify.Controller;


import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Service.CustomUserDetailService;
import com.example.xcelify.Service.UserService;
import lombok.RequiredArgsConstructor;
import com.example.xcelify.Model.User;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.util.Comparator;
import java.util.List;

@Controller
@RequiredArgsConstructor
public class UserController {

    private final CustomUserDetailService customUserDetailService;
    private final UserService userService;
    private final ProductRepository productRepository;
    @GetMapping("/login")
    public String login() {
        return "login";
    }

    @GetMapping("/")
    public String main() {
        return "main";
    }

    @GetMapping("/products")
    public String getProducts(Model model) {
        List<Product> productList = userService.getAllUserProducts();

        productList.sort(Comparator.comparing(Product::getName));
        model.addAttribute("products", productList);

        return "user_products";
    }

    @GetMapping("/documentation")
    public String getDocumentation() {
        return "documentationUser";
    }

    @GetMapping("/settings")
    public String getSettings(Model model) {
        User current = customUserDetailService.getCurrentUser();
        model.addAttribute("user", current);
        return "settings";
    }

    @PostMapping("/products/delete")
    public String deleteProduct(@RequestParam Long id) {
        userService.deleteProduct(id);
        return "redirect:/products";
    }

    @PostMapping("/settings/changepassword")
    public String changePassword(@RequestParam String newPassword) {
        userService.setPassword(newPassword);
        return "redirect:/settings";
    }
}
