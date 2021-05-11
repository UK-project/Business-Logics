package com.eitech1.chartv.controller;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.service.ExcelService;

@RestController
@RequestMapping("/excel")
public class ExcelController {
	
	@Autowired
	private ExcelService excelService;
	
	@PostMapping("/upload")
	public String uploadData(MultipartFile multipartFile) throws EncryptedDocumentException, InvalidFormatException, IOException, ChartVException {
		
			excelService.readExcel(multipartFile);
		
		return null;
		
	}

}
