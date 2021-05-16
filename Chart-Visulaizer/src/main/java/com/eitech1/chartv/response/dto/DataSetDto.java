package com.eitech1.chartv.response.dto;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class DataSetDto {
	
	@JsonProperty("data_set_id")
	private int dataSetId;
	
    @JsonProperty("json_data")
	private String jsonData;



}
