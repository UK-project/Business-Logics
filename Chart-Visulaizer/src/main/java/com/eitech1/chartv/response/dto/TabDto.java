package com.eitech1.chartv.response.dto;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class TabDto {
	
	@JsonProperty("tab_id")
	private int tabId;
	
	@JsonProperty("tab_name")
	private String tabName;
	
	@JsonProperty("tab_topic")
	private String tabTopic;
	
	@JsonProperty("data_set")
	private List<DataSetDto> data;
	

}
