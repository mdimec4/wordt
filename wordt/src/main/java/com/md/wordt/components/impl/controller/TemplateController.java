package com.md.wordt.components.impl.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.tinylog.Logger;

import com.md.wordt.components.interfaces.services.ITemplateService;
import com.md.wordt.dto.ContentControlStructureWrapDTO;
import com.md.wordt.dto.ContentControlValuesWrapDTO;
import com.md.wordt.dto.ErrorDTO;
import com.md.wordt.dto.GeneratedDocumentHolderInternalDTO;
import com.md.wordt.exception.ValidationException;

@RestController
@RequestMapping("docx_template")
public class TemplateController {

	@Autowired
	private ITemplateService templateService;

	@RequestMapping(value = "/upload", method = RequestMethod.POST, consumes = {
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, produces = {
					MediaType.APPLICATION_JSON_VALUE })
	public ResponseEntity<?> getDocumentStructure(@RequestParam("file") MultipartFile document) {
		try {
			ContentControlStructureWrapDTO structDto = templateService.getDocumentStructure(document);
			return new ResponseEntity<>(structDto, HttpStatus.OK);
		} catch (ValidationException e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(e.getMessage());
			Logger.error(e);
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		} catch (Exception e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(ErrorDTO.InternalServerError);
			Logger.warn(e.getMessage());
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		}
	}

	@RequestMapping(value = "/populate", method = RequestMethod.POST, produces = {
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, consumes = {
					MediaType.APPLICATION_JSON_VALUE })
	public ResponseEntity<?> populateDocument(@RequestBody ContentControlValuesWrapDTO values) {
		try {
			GeneratedDocumentHolderInternalDTO internalDto = templateService.populateDocument(values);
			HttpHeaders responseHeaders = new HttpHeaders();
			responseHeaders.set("Content-Disposition",
					String.format("attachment; filename=\"%s\"", internalDto.getFilename()));

			return ResponseEntity.ok().headers(responseHeaders).body(internalDto.getDocumentData());
		} catch (ValidationException e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(e.getMessage());
			Logger.error(e);
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		} catch (Exception e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(ErrorDTO.InternalServerError);
			Logger.warn(e.getMessage());
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		}
	}

	@RequestMapping(value = "/abort", method = RequestMethod.DELETE)
	public ResponseEntity<?> abortAndDelete(@RequestParam("token") String documentRayToken) {
		try {
			templateService.abortAndDelete(documentRayToken);
			return ResponseEntity.ok().build();
		} catch (ValidationException e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(e.getMessage());
			Logger.error(e);
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		} catch (Exception e) {
			ErrorDTO errDto = new ErrorDTO();
			errDto.setError(ErrorDTO.InternalServerError);
			Logger.warn(e.getMessage());
			return new ResponseEntity<>(errDto, HttpStatus.BAD_REQUEST);
		}
	}

}
