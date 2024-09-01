package com.md.wordt.components.impl.services;

import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.RandomStringUtils;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.LoadFromZipNG.ByteArray;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.tinylog.Logger;

import com.md.wordt.components.interfaces.services.ITemplateService;
import com.md.wordt.dto.ContentControlStructureWrapDTO;
import com.md.wordt.dto.ContentControlValuesWrapDTO;
import com.md.wordt.dto.GeneratedDocumentHolderInternalDTO;
import com.md.wordt.dto.util.ContentControlStructureDTO;
import com.md.wordt.exception.InternalException;
import com.md.wordt.exception.ValidationException;
import com.md.wordt.util.WordContentControlAnalayzerUtil;
import com.md.wordt.util.WordContentControlPopulatorUtil;

@Service
public class TemplateService implements ITemplateService {

	private Map<String, StoredDocument> documentStorage = new HashMap<>();

	// For web page see:
	// https://www.freecodecamp.org/news/upload-files-with-javascript/
	@Override
	public ContentControlStructureWrapDTO getDocumentStructure(MultipartFile mFile) {
		if (mFile == null) {
			Logger.error("MultipartFile == null");
			throw new InternalException("MultipartFile == null");
		}
		try {
			WordprocessingMLPackage wordPackageSource = Docx4J.load(mFile.getInputStream());

			WordContentControlAnalayzerUtil wcca = new WordContentControlAnalayzerUtil(wordPackageSource);
			List<ContentControlStructureDTO> structureDTOs = wcca.scanDocument();

			String documentRayToken = generateNewDocumentRayToken();

			ContentControlStructureWrapDTO contentControlStructureWrapDTO = new ContentControlStructureWrapDTO();
			contentControlStructureWrapDTO.setDocumentRayToken(documentRayToken);
			contentControlStructureWrapDTO.setStructure(structureDTOs);

			StoredDocument storedDocument = new StoredDocument();
			storedDocument.document = wordPackageSource;
			storedDocument.fileName = getFilenameFromPath(mFile.getOriginalFilename());
			storedDocument.touch();
			synchronized (documentStorage) {
				documentStorage.put(documentRayToken, storedDocument);
			}

			return contentControlStructureWrapDTO;

		} catch (Docx4JException | IOException e) {
			Logger.error(e);
			throw new ValidationException("File is not valid!");
		}
	}

	// Some browsers like Opera might return whole path for file
	private String getFilenameFromPath(String path) {
		int position = path.lastIndexOf("/");
		if (position == -1) {
			position = path.lastIndexOf("\\");
		}

		if (position == -1) {
			return path;
		}
		return path.substring(position + 1);
	}

	@Override
	public GeneratedDocumentHolderInternalDTO populateDocument(ContentControlValuesWrapDTO values) {
		StoredDocument storedDocument = null;
		synchronized (documentStorage) {
			storedDocument = documentStorage.get(values.getDocumentRayToken());
			if (storedDocument != null) {
				if (storedDocument.isExpired()) {
					documentStorage.remove(values.getDocumentRayToken());
					throw new ValidationException("Document has expired!");
				}
				storedDocument.touch();
			}
		}

		if (storedDocument == null) {
			throw new ValidationException("Document not found!");
		}
		WordprocessingMLPackage documentClone = (WordprocessingMLPackage) storedDocument.document.clone();

		WordContentControlPopulatorUtil ccPopulator = new WordContentControlPopulatorUtil(documentClone);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
			ccPopulator.populateDocument(values.getValues());
			ccPopulator.save(bos);
		} catch (Docx4JException e) {
			Logger.error(e);
			throw new InternalException("Error while populating or generating a document!");
		}
		
		String outFileName = generateOutputFileName(storedDocument.fileName);

		// TODO: pack result in DTO
		return null;
	}

	@Override
	public void abortAndDelete(String documentRayToken) {
		synchronized (documentStorage) {
			documentStorage.remove(documentRayToken);
		}

	}

	@Scheduled(initialDelay = 3600, fixedRate = 3600, timeUnit = TimeUnit.SECONDS)
	public void documentCleaner() {
		List<String> toRemove = new ArrayList<String>();
		synchronized (documentStorage) {
			for (String documentRay : documentStorage.keySet()) {
				StoredDocument storedDocument = documentStorage.get(documentRay);
				if (storedDocument.isExpired()) {
					toRemove.add(documentRay);
				}
			}

			for (String documentRay : toRemove) {
				documentStorage.remove(documentRay);
			}

		}

	}

	private String generateNewDocumentRayToken() {
		String generatedString = RandomStringUtils.random(21, true, true);

		Instant now = Instant.now();
		long nano = (now.getEpochSecond() * 1_000_000_000) + now.getNano();

		return generatedString + Long.toHexString(nano);
	}
	
	private String generateOutputFileName(String fileName) {
		if(fileName == null || fileName.isEmpty())
			return "document.docx";
		
		String generatedString = RandomStringUtils.random(4, true, true);
		Instant now = Instant.now();
		long nano = (now.getEpochSecond() * 1_000_000_000) + now.getNano();
		String unique = generatedString + Long.toHexString(nano);
		
		int index = fileName.lastIndexOf(".");
		
		return fileName.substring(0, index -1) + "_" + unique + ".docx";
	}

	private class StoredDocument {
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
