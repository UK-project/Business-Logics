package com.eitech1.chartv.exceptions.handler;

import java.util.Map;
import java.util.stream.Collectors;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.FieldError;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.context.request.WebRequest;
import org.springframework.web.servlet.mvc.method.annotation.ResponseEntityExceptionHandler;

import com.eitech1.chartv.exceptions.ChartVDataValidationException;
import com.eitech1.chartv.exceptions.ChartVException;
import com.eitech1.chartv.exceptions.ChartVPersistenceException;
import com.eitech1.chartv.response.template.Response;

@ControllerAdvice
public class ControllerExceptionHandler extends ResponseEntityExceptionHandler {
	
	//private static Logger Log = AppLogger.getLogger();
	private static final String LOG_MESSAGE = "Exception caught and handled.";
	
	@ExceptionHandler({ ChartVPersistenceException.class })
	public ResponseEntity<Response<Object>> handleChartVPersistenceException(ChartVPersistenceException ex) {
		//Log.error(LOG_MESSAGE, ex);
		String message = ex.getMessage();
		if(message == null) {
			message = ex.getCause().getMessage();
		}
		return Response.error(null, message, HttpStatus.BAD_REQUEST);
	}
	
	@ExceptionHandler({ ChartVDataValidationException.class })
	public ResponseEntity<Response<Object>> handlChartVDataValidationException(ChartVDataValidationException ex) {
		//Log.error(LOG_MESSAGE, ex);
		String message = ex.getMessage();
		if(message == null && ex.getCause() != null) {
			message = ex.getCause().getMessage();
		}
		return Response.error(null, message, HttpStatus.BAD_REQUEST);
	}

	@ExceptionHandler({ ChartVException.class })
	public ResponseEntity<Response<Object>> handlChartVException(ChartVException ex) {
		//Log.error(LOG_MESSAGE, ex);
		String message = ex.getMessage();
		if(message == null && ex.getCause() != null) {
			message = ex.getCause().getMessage();
		}
		return Response.error(null, message, HttpStatus.INTERNAL_SERVER_ERROR);
	}
	
	@Override
	public ResponseEntity<Object> handleMethodArgumentNotValid(MethodArgumentNotValidException ex, HttpHeaders headers, HttpStatus status, WebRequest request) {
	//	Log.error(LOG_MESSAGE, ex);
		Map<String, String> result = ex.getBindingResult().getFieldErrors().stream().collect(Collectors.toMap(FieldError::getField, FieldError::getDefaultMessage));
		return new ResponseEntity<>(Response.error(result, HttpStatus.BAD_REQUEST.toString(), HttpStatus.BAD_REQUEST).getBody(), null,HttpStatus.BAD_REQUEST);
	}


}
