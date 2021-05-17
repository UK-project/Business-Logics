package com.eitech1.chartv.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.response.dto.SheetMetaDto;
import com.eitech1.chartv.response.template.Response;
import com.eitech1.chartv.service.MetaService;

import io.swagger.annotations.ApiOperation;



/**
 * @author Nilar
 *
 */
@RestController
@RequestMapping("/meta")
@CrossOrigin(origins = "*")
public class MetaController {
	@Autowired
	private MetaService metaService;
	
	
	@GetMapping("/sheet")
	@ApiOperation("get all master sheet details")
	public ResponseEntity<Response<List<SheetMetaDto>>> getAllSheetMetas() throws ChartVException{
		return metaService.getAllSheetMeta();
		
	}

}
