package com.eitech1.chartv.response.dto;

import java.util.Date;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class SheetExDto {
	
	@JsonProperty("sheet_id")
	private int sheetId;

	@JsonProperty("uploaded_date")
	private Date uploadedDate;
	
	@JsonProperty("tabs")
	private List<TabDto> tabs;
	
	@JsonProperty("sheet_meta")
	private SheetMetaDto sheetMeta;
	

}
