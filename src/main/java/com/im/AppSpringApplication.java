package com.im;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


@MapperScan("com.im.modules.dao")
@SpringBootApplication
public class AppSpringApplication {
    public static void main(String[] args) {
        SpringApplication.run(AppSpringApplication.class,args);
    }
}