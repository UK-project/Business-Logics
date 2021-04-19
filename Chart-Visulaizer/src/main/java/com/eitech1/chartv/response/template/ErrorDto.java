package com.eitech1.chartv.response.template;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonPropertyOrder;

import lombok.Builder;
import lombok.Data;

import com.fasterxml.jackson.annotation.JsonInclude.Include;




@Data
@Builder
@JsonPropertyOrder(alphabetic = true)
public class ErrorDto {
	public static final int GENERAL_ERROR = 1;
	@JsonInclude(Include.NON_NULL)
	@JsonProperty("error_code")
	private Integer errorCode;
	@JsonInclude(Include.NON_NULL)
	@JsonProperty("error_message")
	private String message;
}
