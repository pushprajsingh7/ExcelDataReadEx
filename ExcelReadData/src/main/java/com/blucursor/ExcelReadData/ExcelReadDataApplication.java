package com.blucursor.ExcelReadData;

import java.io.IOException;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelReadDataApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ExcelReadDataApplication.class, args);
		ExcelKRI ex = new ExcelKRI();
		ex.excel();
	}

}
