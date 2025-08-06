package com.jnj.extracter;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExtracterApplication {

	public static void main(String[] args) {
		ZipSecureFile.setMinInflateRatio(0.001);
		
		SpringApplication.run(ExtracterApplication.class, args);
	}

}
