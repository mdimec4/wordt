package com.md.wordt.dto;

import java.util.Collections;
import java.util.List;

import com.md.wordt.dto.util.ContentControlStructureDTO;

public class ContentControlStructureWrapDTO {
	private String documentRayToken;
	private String error;

	private List<ContentControlStructureDTO> structure = Collections.emptyList();

	public String getDocumentRayToken() {
		return documentRayToken;
	}

	public void setDocumentRayToken(String documentRayToken) {
		this.documentRayToken = documentRayToken;
	}

	public String getError() {
		return error;
	}

	public void setError(String error) {
		this.error = error;
	}

	public List<ContentControlStructureDTO> getStructure() {
		return structure;
	}

	public void setStructure(List<ContentControlStructureDTO> structure) {
		this.structure = structure;
	}
}
