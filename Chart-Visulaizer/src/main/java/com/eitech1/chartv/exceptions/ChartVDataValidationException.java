package com.eitech1.chartv.exceptions;



/**
 * Use this exception class to throw the data validation exceptions such as invalid role code, duplicate username etc
 * 
 * @author Nilar
 *
 */
public class ChartVDataValidationException extends ChartVException {

	private static final long serialVersionUID = 1L;

	public ChartVDataValidationException(String message, Object invalidValue) {
		super(message + " [" + invalidValue.toString() + "]");
		
	}

	public ChartVDataValidationException(String message) {
		super(message);
	}

}
