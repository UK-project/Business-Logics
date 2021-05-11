package com.eitech1.chartv.util;

import java.io.File;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Component;

import com.eitech1.chartv.exceptions.ChartVException;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

@Component
public class ExcelUtil {

	public Workbook getExcel(String excelPath) throws ChartVException {

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(new File(excelPath));
		} catch (EncryptedDocumentException e) {
			throw new ChartVException("The uploaded file is encrypted", e.getCause());
		} catch (InvalidFormatException e) {
			throw new ChartVException("The uploaded file format is incorrect", e.getCause());
		} catch (Exception e) {
			throw new ChartVException("Unexpected error occured", e.getCause());
		}
//				
//		// Retrieving the number of sheets in the Workbook
//		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
//
//		//Retrieving Sheets using for-each loop
//		for (Sheet sheet : workbook) {
//			System.out.println("=> " + sheet.getSheetName());
//		}
//		
		return workbook;
	}

	public String createRowJson(List<String> excelData, List<String> headerList) throws ChartVException {

		// TelecomAdSpend telAdSpend = new TelecomAdSpend();
		HashMap<String, String> map = new HashMap<String, String>();

		int i = 0;
		for (String cellvalue : excelData) {
			map.put(headerList.get(i), cellvalue);
			i++;
		}

		String json;
		try {
			json = new ObjectMapper().writeValueAsString(map);
		} catch (JsonProcessingException e) {
			throw new ChartVException(e.getCause());
		}
		// telAdSpend.setJsonString(json);
		System.out.println(json);

		return json;
	}

}
