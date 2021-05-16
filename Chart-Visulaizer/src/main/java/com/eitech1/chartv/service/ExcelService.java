package com.eitech1.chartv.service;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import com.eitech1.chartv.Entity.SheetEx;
import com.eitech1.chartv.Entity.Tab;
import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.response.dto.SheetExDto;
import com.eitech1.chartv.response.template.Response;


public interface ExcelService {
	
	ResponseEntity<Response<SheetExDto>> readExcel(MultipartFile multipartFile) throws EncryptedDocumentException, InvalidFormatException, IOException, ChartVException;

}
