package com.eitech1.chartv.service.util;

import java.util.Date;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.eitech1.chartv.Entity.DataSet;
import com.eitech1.chartv.Entity.SheetEx;
import com.eitech1.chartv.Entity.Tab;
import com.eitech1.chartv.respository.SheetMetaRepository;

@Component
public class DtoToEntityMapper {
	
	@Autowired
	private SheetMetaRepository sheetMetaRepository;
	
	public DataSet convertToDataSet(String jsonData,Tab tab) {
		return DataSet.builder()
				.jsonData(jsonData)
				.tab(tab)
				.build();
		
	}
	
	public Tab convertToTab(String tabName,SheetEx sheet) {
		return Tab.builder()
				.tabName(tabName)
				.sheet(sheet)
				.build();
	}
	public SheetEx converToSheet() {
		return SheetEx.builder()
				.uploadedDate(new Date())
				.sheetMeta(sheetMetaRepository.getOne(1))
				.build();
		
	}
	
	

}
