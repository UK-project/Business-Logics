package com.eitech1.chartv.service.impl;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.eitech1.chartv.Entity.SheetMeta;
import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.messages.ResponseMessages;
import com.eitech1.chartv.response.dto.SheetMetaDto;
import com.eitech1.chartv.response.template.Response;
import com.eitech1.chartv.respository.SheetMetaRepository;
import com.eitech1.chartv.service.util.EntityToDtoMapper;

@Service
public class MetaService implements com.eitech1.chartv.service.MetaService {
	@Autowired
	private SheetMetaRepository sheetMetaRepository;

	@Autowired
	private EntityToDtoMapper entityToDtoMapper;

	@Override
	public ResponseEntity<Response<List<SheetMetaDto>>> getAllSheetMeta() throws ChartVException {

		try {
			List<SheetMeta> sheetMeta = sheetMetaRepository.findAll();
			List<SheetMetaDto> sheetMetaDto = new ArrayList<SheetMetaDto>();
			for (SheetMeta sheetM : sheetMeta) {
				sheetMetaDto.add(entityToDtoMapper.convertToSheetMetaDto(sheetM));
				
			}
			return Response.success(sheetMetaDto, HttpStatus.OK);
			
		} catch (Exception e) {
			throw new ChartVException(ResponseMessages.UNEXPECTED_ERROR, e.getCause());
		}

		
	}

}
