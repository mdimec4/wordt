package com.md.wordt.components.impl.controller;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.tinylog.Logger;

import com.md.wordt.dto.ErrorDTO;
import com.md.wordt.exception.InternalException;
import com.md.wordt.exception.ValidationException;

@ControllerAdvice
public class MdControllerAdvice {

	@ExceptionHandler(ValidationException.class)
	ResponseEntity<ErrorDTO> handleValidationException(ValidationException e) {
		ErrorDTO errDto = new ErrorDTO();
		errDto.setError(e.getMessage());
		Logger.warn(e.getMessage());
		return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
	}
	
	@ExceptionHandler(Exception.class)
	ResponseEntity<ErrorDTO> handleException(InternalException e) {
		ErrorDTO errDto = new ErrorDTO();
		errDto.setError(ErrorDTO.InternalServerError);
		Logger.error(e);
		return new ResponseEntity<>(errDto, HttpStatus.INTERNAL_SERVER_ERROR);
	}

}
