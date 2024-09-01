package com.md.wordt.components.interfaces.services;

import org.springframework.web.multipart.MultipartFile;

import com.md.wordt.dto.ContentControlStructureWrapDTO;
import com.md.wordt.dto.ContentControlValuesWrapDTO;
import com.md.wordt.dto.GeneratedDocumentHolderInternalDTO;

public interface ITemplateService {
	ContentControlStructureWrapDTO getDocumentStructure(MultipartFile document);

	GeneratedDocumentHolderInternalDTO populateDocument(ContentControlValuesWrapDTO values);

	void abortAndDelete(String documentRayToken);
}
