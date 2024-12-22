package com.example.xcelify.Service;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Repository.ProductRepository;
import com.example.xcelify.Repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import com.example.xcelify.Model.User;

@Service
@Slf4j
@RequiredArgsConstructor
public class UserService {

    private final ProductRepository productRepository;
    private final CustomUserDetailService customUserDetailService;
    private final UserRepository userRepository;
    public List<Product> getAllUserProducts() {

        User currentUser = customUserDetailService.getCurrentUser();
        return productRepository.findAllByUser(currentUser);
    }

    public void setPassword(String password) {
        User currentUser = customUserDetailService.getCurrentUser();

        if (password.length() < 8) {
            throw new IllegalArgumentException("Пароль должен содержать хотя бы 8 символов.");
        }

        currentUser.setPassword(password);
        userRepository.save(currentUser);
    }

    public void deleteProduct(Long id) {
        productRepository.deleteById(id);
    }

}
