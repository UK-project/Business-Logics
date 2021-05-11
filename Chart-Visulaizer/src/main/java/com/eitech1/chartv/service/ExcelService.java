package com.eitech1.chartv.service;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import com.eitech1.chartv.Entity.SheetEx;
import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.response.template.Response;


public interface ExcelService {
	
	ResponseEntity<Response<SheetEx>> readExcel(MultipartFile multipartFile) throws EncryptedDocumentException, InvalidFormatException, IOException, ChartVException;

}
