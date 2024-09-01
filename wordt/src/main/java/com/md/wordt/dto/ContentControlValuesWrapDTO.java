package com.md.wordt.dto;

import java.util.Collections;
import java.util.List;
import java.util.Map;

import com.md.wordt.dto.util.ContentControlValuesDTO;

public class ContentControlValuesWrapDTO {
	private String documentRayToken;
	private Map<String, List<ContentControlValuesDTO>> values = Collections.emptyMap();

	public String getDocumentRayToken() {
		return documentRayToken;
	}

	public void setDocumentRayToken(String documentRayToken) {
		this.documentRayToken = documentRayToken;
	}

	public Map<String, List<ContentControlValuesDTO>> getValues() {
		return values;
	}

	public void setValues(Map<String, List<ContentControlValuesDTO>> values) {
		this.values = values;
	}

}
