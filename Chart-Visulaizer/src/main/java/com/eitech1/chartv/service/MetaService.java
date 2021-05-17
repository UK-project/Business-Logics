package com.eitech1.chartv.service;

import java.util.List;

import org.springframework.http.ResponseEntity;

import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.response.dto.SheetMetaDto;
import com.eitech1.chartv.response.template.Response;

public interface MetaService  {
	
	public ResponseEntity<Response<List<SheetMetaDto>>>  getAllSheetMeta() throws ChartVException;

}
