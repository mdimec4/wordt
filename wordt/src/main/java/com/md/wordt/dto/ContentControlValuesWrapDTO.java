package com.md.wordt.dto;

import java.util.Collections;
import java.util.List;

public class ContentControlValuesWrapDTO {
	private String documentRayToken;
	private List<ContentControlValuesWrapDTO> values = Collections.emptyList();

	public String getDocumentRayToken() {
		return documentRayToken;
	}

	public void setDocumentRayToken(String documentRayToken) {
		this.documentRayToken = documentRayToken;
	}

	public List<ContentControlValuesWrapDTO> getValues() {
		return values;
	}

	public void setValues(List<ContentControlValuesWrapDTO> values) {
		this.values = values;
	}

}
