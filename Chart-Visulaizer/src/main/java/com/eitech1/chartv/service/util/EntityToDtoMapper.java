package com.eitech1.chartv.service.util;

import java.util.List;

import org.springframework.stereotype.Component;

import com.eitech1.chartv.Entity.DataSet;
import com.eitech1.chartv.Entity.SheetEx;
import com.eitech1.chartv.Entity.SheetMeta;
import com.eitech1.chartv.Entity.Tab;
import com.eitech1.chartv.response.dto.DataSetDto;
import com.eitech1.chartv.response.dto.SheetExDto;
import com.eitech1.chartv.response.dto.SheetMetaDto;
import com.eitech1.chartv.response.dto.TabDto;

@Component
public class EntityToDtoMapper {
	
	public DataSetDto convertToDataSetDto(DataSet dataset) {
		return DataSetDto.builder()
				.dataSetId(dataset.getDataSetId())
				.jsonData(dataset.getJsonData())
				.build();
		
	}
	
	public TabDto convertToTabDto(Tab tab,List<DataSetDto> dataset) {
		return TabDto.builder()
				.tabId(tab.getTabId())
				.tabName(tab.getTabName())
				.tabTopic(tab.getTabTopic())
				.data(dataset)
				.build();
			
	}
	public SheetExDto converToSheetExDto(SheetEx sheet,List<TabDto> tabs) {
	SheetMetaDto sheetMetaDto =	SheetMetaDto.builder().id(sheet.getSheetMeta().getId())
		.sheetCode(sheet.getSheetMeta().getSheetCode())
		.sheetName(sheet.getSheetMeta().getSheetName())
		.build();
		return SheetExDto.builder()
				.sheetId(sheet.getSheetId())
				.sheetMeta(sheetMetaDto)
				.uploadedDate(sheet.getUploadedDate())
				.tabs(tabs)
				.build();
	}
	
	public SheetMetaDto convertToSheetMetaDto(SheetMeta sheetMeta) {
		return SheetMetaDto.builder()
				.id(sheetMeta.getId())
				.sheetCode(sheetMeta.getSheetCode())
				.sheetName(sheetMeta.getSheetName())
				.build();
			
	}

}
