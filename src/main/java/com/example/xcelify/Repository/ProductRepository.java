package com.example.xcelify.Repository;

import com.example.xcelify.Model.Product;
import com.example.xcelify.Model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;

@Repository
public interface ProductRepository extends JpaRepository<Product, Long> {
        @Query("SELECT p.cost FROM Product p WHERE p.articul = :articul ORDER BY p.updateCost DESC")
        Double findCostByArticul(@Param("articul") String articul);

        Product findByArticulAndUser(String normalizedArticul, User currentUser);

        @Query("SELECT p.cost FROM Product p WHERE p.articul = :articul AND p.user = :user")
        Double findCostByArticulAndUser(@Param("articul") String articul, @Param("user") User user);

        List<Product> findAllByUser(User user);


}
