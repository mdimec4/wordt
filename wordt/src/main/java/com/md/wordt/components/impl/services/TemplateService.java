package com.md.wordt.components.impl.services;

import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.md.wordt.components.interfaces.services.ITemplateService;
import com.md.wordt.dto.ContentControlStructureWrapDTO;
import com.md.wordt.dto.ContentControlValuesWrapDTO;
import com.md.wordt.dto.GeneratedDocumentHolderInternalDTO;

@Service
public class TemplateService implements ITemplateService {

	private Map<String, DocumentStore> documentStorage = new HashMap<>();

	// For web page see:
	// https://www.freecodecamp.org/news/upload-files-with-javascript/
	@Override
	public ContentControlStructureWrapDTO getDocumentStructure(MultipartFile document) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public GeneratedDocumentHolderInternalDTO populateDocument(ContentControlValuesWrapDTO values) {
		// TODO Auto-generated method stub
		return null;
	}
	
	@Scheduled(initialDelay = 3600, fixedRate = 3600, timeUnit = TimeUnit.SECONDS)
	public void documentCleaner() {
		List<String> toRemove = new ArrayList<String>();
		synchronized (documentStorage) {
			for (String documentRay : documentStorage.keySet()) {
				DocumentStore storedDocument = documentStorage.get(documentRay);
				if (storedDocument.isExpired()) {
					toRemove.add(documentRay);
				}
			}
			
			for (String documentRay : toRemove) {
				documentStorage.remove(documentRay);
			}
			
		}
		
	}

	private class DocumentStore {
		public WordprocessingMLPackage document;
		public String fileName;
		public Instant expires;

		public void touch() {
			expires = Instant.now().plusSeconds(3600);
		}
		
		public boolean isExpired() {
			if (expires == null)
				return true;
			return !Instant.now().isBefore(expires);
		}
	}
}
