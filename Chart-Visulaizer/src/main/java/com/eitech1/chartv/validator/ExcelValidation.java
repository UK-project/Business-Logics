package com.eitech1.chartv.validator;

import java.util.List;

import org.springframework.stereotype.Component;

import com.eitech1.chartv.exceptions.ChartVDataValidationException;

@Component
public class ExcelValidation {
	
	String sheet="Telecoms Adspend Forecasts.xlsx";

	public void validateColumn(List<String> headerList, List<String> rowValueList) throws ChartVDataValidationException  {

		if (headerList.size() != rowValueList.size()) {
			throw new ChartVDataValidationException(
					"length of header columns and length of row columns are different " , headerList + " " + rowValueList);
		}

	}
	
	public void validateSheetName(String sheetName) throws ChartVDataValidationException   {

		if (sheetName.equals(sheet) == false ) {
			throw new ChartVDataValidationException("Sheet name should follow the structure", sheetName);
		}

	}

}
