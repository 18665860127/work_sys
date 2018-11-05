package com.jiang.work_sys;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootApplication(exclude = DataSourceAutoConfiguration.class)
public class WorkSysApplication {

	public static void main(String[] args) {
		SpringApplication.run(WorkSysApplication.class, args);
	}
}
