package com.eitech1.chartv.response.template;



import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Builder
public class MetaDto {
	/**
	 * The success/error status of the API request. Success-1, Error-0
	 */
	
	private Integer status;
	private String detail;

}
