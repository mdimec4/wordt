package com.md.wordt.dto;

public class ErrorDTO {
	public static final String InternalServerError = "Internal server error!";
	private String error;

	public String getError() {
		return error;
	}

	public void setError(String error) {
		this.error = error;
	}
	
}
