package com.eitech1.chartv.exceptions;



/**
 * This is the super class of all the custom exceptions created in the project
 * 
 * @author Nilar
 *
 */
public class ChartVException extends Exception {

	private static final long serialVersionUID = 1L;

	public ChartVException(String message, Throwable cause) {
		super(message, cause);
		// TODO Auto-generated constructor stub
	}

	public ChartVException(String message) {
		super(message);
		// TODO Auto-generated constructor stub
	}

	public ChartVException(Throwable cause) {
		super(cause);
		// TODO Auto-generated constructor stub
	}
	
	

}
