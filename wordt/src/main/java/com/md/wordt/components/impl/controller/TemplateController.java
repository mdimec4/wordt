package com.md.wordt.components.impl.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.md.wordt.components.interfaces.services.ITemplateService;
import com.md.wordt.dto.ContentControlStructureWrapDTO;
import com.md.wordt.dto.ContentControlValuesWrapDTO;
import com.md.wordt.dto.GeneratedDocumentHolderInternalDTO;

@RestController
@RequestMapping("docx_template")
public class TemplateController {

	@Autowired
	private ITemplateService templateService;

	@RequestMapping(value = "/upload", method = RequestMethod.POST, consumes = {
			MediaType.MULTIPART_FORM_DATA_VALUE }, produces = { MediaType.APPLICATION_JSON_VALUE })
	public ContentControlStructureWrapDTO getDocumentStructure(@RequestParam("file") MultipartFile document) {
		return templateService.getDocumentStructure(document);
	}

	@RequestMapping(value = "/populate", method = RequestMethod.POST, produces = {
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, consumes = {
					MediaType.APPLICATION_JSON_VALUE })
	public ResponseEntity<byte[]> populateDocument(@RequestBody ContentControlValuesWrapDTO values) {
		GeneratedDocumentHolderInternalDTO internalDto = templateService.populateDocument(values);
		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Disposition",
				String.format("attachment; filename=\"%s\"", internalDto.getFilename()));

		return ResponseEntity.ok().headers(responseHeaders).body(internalDto.getDocumentData());

	}

	@RequestMapping(value = "/abort", method = RequestMethod.DELETE)
	public void abortAndDelete(@RequestParam("token") String documentRayToken) {
		templateService.abortAndDelete(documentRayToken);
	}

}
