package com.eitech1.chartv.response.dto;

import javax.persistence.Id;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class SheetMetaDto {
	
	@Id
	private int id;
	
	@JsonProperty("sheet_name")
	private String sheetName;
	
	@JsonProperty("sheet_code")
	private String sheetCode;

}
