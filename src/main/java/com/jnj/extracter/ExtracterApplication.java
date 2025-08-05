package com.jnj.extracter;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExtracterApplication {

	public static void main(String[] args) {
		// Disable zip bomb detection by setting a very low minimum inflate ratio
		// This allows processing of Excel files with unusual compression ratios
		ZipSecureFile.setMinInflateRatio(0.001);
		
		SpringApplication.run(ExtracterApplication.class, args);
	}

}
