package com.eitech1.chartv.service.impl;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.hibernate.exception.ConstraintViolationException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.eitech1.chartv.Entity.DataSet;
import com.eitech1.chartv.Entity.SheetEx;
import com.eitech1.chartv.Entity.Tab;
import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.exceptions.ChartVPersistenceException;
import com.eitech1.chartv.response.template.Response;
import com.eitech1.chartv.respository.SheetExRepository;
import com.eitech1.chartv.service.ExcelService;
import com.eitech1.chartv.service.util.DtoToEntityMapper;
import com.eitech1.chartv.util.ExcelUtil;

@Service
public class ExcelServiceImpl implements ExcelService{
	
	@Autowired 
	private ExcelUtil excelUtil;
	
	@Autowired
	private SheetExRepository sheetExRepository;
	
	@Autowired
	private DtoToEntityMapper dtoDtoToEntityMapper;
	
	private String excelPath ="D:\\test excel\\";

	@Override
	public ResponseEntity<Response<SheetEx>> readExcel(MultipartFile multipartFile) throws  ChartVException {
		
		try {
			
		String filepath = excelPath + multipartFile.getOriginalFilename();
		
		 byte[]	bytes = multipartFile.getBytes();
			java.nio.file.Path path = Paths.get(excelPath + multipartFile.getOriginalFilename());
			Files.write(path, bytes);
		
		Workbook workbook=excelUtil.getExcel(filepath);
		int x = workbook.getNumberOfSheets();
		
		String jsonData=null; //dataset json data initialize
		List<DataSet> dataSetList=new ArrayList<DataSet>(); //dataset list to set with tab
		List<Tab> tabList=new ArrayList<Tab>();//tab list to set with sheetEx
		String tabTopic=null;//tab topic
		
		//dto to entity mapping (Sheet)
		 SheetEx sheetEx = dtoDtoToEntityMapper.converToSheet();
		
		// Getting the Sheet at index zero
		for (int i = 0; i < x; i++) {
			String header=null;
			List<String> headerList=new ArrayList<String>();

			Sheet sheetTab = workbook.getSheetAt(i);
			
			//dto to entity mapping (Tab)
			Tab tab= dtoDtoToEntityMapper.convertToTab(sheetTab.getSheetName(), sheetEx);
			
			//int noOfColumns = sheetTab.getRow(2).getLastCellNum();

			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			System.out.println("The sheet number is " + i + 1);
			// 2. Or you can use a for-each loop to iterate over the rows and columns
			System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");

			for (Row row : sheetTab) {
			
				List<String> rowValueList = new ArrayList<String>();
				
//				System.out.println(row.getRowNum());
//				if (row.getRowNum() == 0) {
//					tabTopic=
//					continue; // just skip the rows if row number is 0 or 1
//				}
 
				for (Cell cell : row) {
					cell.setCellType(CellType.STRING);
					
					if (row.getRowNum() == 0) {
						tabTopic = dataFormatter.formatCellValue(cell);
					}

					if (row.getRowNum() == 1) {
						 header = dataFormatter.formatCellValue(cell);
						 headerList.add(header); System.out.println(header);
					}
					else {
					String cellValue = dataFormatter.formatCellValue(cell);
					System.out.print(cellValue + "\t");
					rowValueList.add(cellValue);
					}

				}
				System.out.println();
						if (row.getRowNum() != 0 && row.getRowNum() != 1 ) {
							jsonData = excelUtil.createRowJson(rowValueList, headerList);
						    DataSet  dataSet = dtoDtoToEntityMapper.convertToDataSet(jsonData, tab);		
						    dataSetList.add(dataSet);
						}
			
			} 
			tab.setData(dataSetList);
			tab.setTabTopic(tabTopic);
			tabList.add(tab);
			
		}
		
		sheetEx.setTabs(tabList);
		sheetExRepository.save(sheetEx);
		
		return Response.success(sheetEx, HttpStatus.OK);

		} catch (ChartVException e) {
			throw e;
		} catch (Exception e) {

			if(e instanceof org.springframework.dao.DataIntegrityViolationException && e.getCause() instanceof org.hibernate.exception.ConstraintViolationException) {
				org.hibernate.exception.ConstraintViolationException hibernateException = (ConstraintViolationException) e.getCause();
				throw new ChartVPersistenceException(hibernateException.getCause().getLocalizedMessage(), hibernateException.getCause());
				
			}
			throw new ChartVException("Unexpected error occured", e);
		}
		
		
	}

}
