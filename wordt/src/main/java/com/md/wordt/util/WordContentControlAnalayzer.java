package com.md.wordt.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.CTSdtCell;
import org.docx4j.wml.CTSdtRow;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtElement;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tag;

import com.md.wordt.dto.ContentControlStructureDTO;

import jakarta.xml.bind.JAXBElement;

public class WordContentControlAnalayzer {

	private WordprocessingMLPackage wordPackage;

	public WordContentControlAnalayzer(WordprocessingMLPackage wordPackage) {
		this.wordPackage = wordPackage;

	}

	public WordContentControlAnalayzer(InputStream inputStream) throws Docx4JException {
		wordPackage = Docx4J.load(inputStream);

	}

	public WordContentControlAnalayzer(File file) throws Docx4JException, FileNotFoundException {
		this(new FileInputStream(file));
	}

	public static String getSdtTag(SdtElement sdtElement) {
		Optional<Object> tagObjOpt = sdtElement.getSdtPr().getRPrOrAliasOrLock() //
				.stream().filter(o -> (o instanceof Tag)) //
				.findFirst();

		if (tagObjOpt.isEmpty())
			return "";
		Tag tag = (Tag) tagObjOpt.get();
		return tag.getVal();
	}

	@SuppressWarnings("unchecked")
	public static String getSdtName(SdtElement sdtElement) {
		Optional<org.docx4j.wml.SdtPr.Alias> aliasOpt = sdtElement.getSdtPr().getRPrOrAliasOrLock() //
				.stream() //
				.filter(o -> {
					if (!(o instanceof JAXBElement<?>))
						return false;
					JAXBElement<?> j = (JAXBElement<?>) o;
					if (!j.getDeclaredType().equals(org.docx4j.wml.SdtPr.Alias.class))
						return false;
					if (!j.getName().getLocalPart().equals("alias"))
						return false;
					return true;

				}) //
				.map(o -> {
					JAXBElement<org.docx4j.wml.SdtPr.Alias> j = (JAXBElement<org.docx4j.wml.SdtPr.Alias>) o;
					return j.getValue();
				}) //
				.findFirst();
		if (aliasOpt.isEmpty())
			return "";
		return aliasOpt.get().getVal();
	}

	@SuppressWarnings("unchecked")
	public static boolean isW15RepeatingSection(SdtElement sdtElement) {
		Optional<JAXBElement<org.docx4j.w15.CTSdtRepeatedSection>> rsOpt = sdtElement.getSdtPr().getRPrOrAliasOrLock() //
				.stream() //
				.filter(o -> {
					if (!(o instanceof JAXBElement<?>))
						return false;
					JAXBElement<?> j = (JAXBElement<?>) o;
					if (!j.getDeclaredType().equals(org.docx4j.w15.CTSdtRepeatedSection.class))
						return false;

					return true;

				})//
				.map(o -> (JAXBElement<org.docx4j.w15.CTSdtRepeatedSection>) o)//
				.findAny();
		return rsOpt.isPresent();
	}

	public static boolean isW15RepeatingSectionItem(SdtElement sdtElement) {
		Optional<Object> rsiOpt = sdtElement.getSdtPr().getRPrOrAliasOrLock() //
				.stream() //
				.filter(o -> {
					if (!(o instanceof JAXBElement<?>))
						return false;
					JAXBElement<?> j = (JAXBElement<?>) o;
					if (!j.getName().getLocalPart().equals("repeatingSectionItem"))
						return false;

					return true;

				})//
				.findAny();
		return rsiOpt.isPresent();
	}

	private void scanDocumentRecursive(List<Object> docElementContent, List<ContentControlStructureDTO> dtoContent) {
		if (docElementContent == null || dtoContent == null)
			return;

		for (Object object : docElementContent) {
			if (object instanceof JAXBElement<?>) {
				object = ((JAXBElement<?>) object).getValue();
			}
			System.out.println(object.getClass().getSimpleName());

			List<Object> childContent = null;
			List<ContentControlStructureDTO> childDTOs = null;
			if (object instanceof SdtElement) {
				SdtElement sdtElement = (SdtElement) object;
				childContent = sdtElement.getSdtContent().getContent();

				String tag = getSdtTag(sdtElement);
				System.out.println("Tag: " + tag);
				String name = getSdtName(sdtElement);
				System.out.println("Name: " + name);
				boolean isRepeatingSection = isW15RepeatingSection(sdtElement);
				System.out.println("isRepeatingSection: " + isRepeatingSection);
				// boolean isRepeatingSectionItem = isW15RepeatingSectionItem(sdtElement);
				// System.out.println("isRepatingSectionItem: " + isRepeatingSectionItem);

				boolean isMultiParagraph = false;
				if (sdtElement instanceof CTSdtCell) {
					isMultiParagraph = false;
				} else if (sdtElement instanceof CTSdtRow) {
					isMultiParagraph = false;
				} else if (sdtElement instanceof SdtBlock) {
					isMultiParagraph = true;
				} else if (sdtElement instanceof SdtRun) {
					isMultiParagraph = false;
				}
				System.out.println("MultiParagraph: " + isMultiParagraph);

				if (!tag.isEmpty()) {
					ContentControlStructureDTO dto = new ContentControlStructureDTO();
					dto.setTag(tag);
					dto.setName(name);
					dto.setMultiParagraph(isMultiParagraph);
					dto.setRepeating(isRepeatingSection);
					dtoContent.add(dto);

					childDTOs = dto.getChildren();
				}

			} else if (object instanceof ContentAccessor) {
				ContentAccessor ca = (ContentAccessor) object;
				childContent = ca.getContent();
			}
			
			if (childContent != null && childDTOs == null)
				scanDocumentRecursive(childContent, dtoContent);
			else if (childContent != null && childDTOs != null)
				scanDocumentRecursive(childContent, childDTOs);
		}
	}

	public List<ContentControlStructureDTO> scanDocument() {
		List<ContentControlStructureDTO> rootList = new ArrayList<ContentControlStructureDTO>();
		wordPackage.getMainDocumentPart();
		scanDocumentRecursive(wordPackage.getMainDocumentPart().getContent(), rootList);

		return rootList;
	}

	public static void main(String[] args) throws FileNotFoundException, Docx4JException {
		WordContentControlAnalayzer wcca = new WordContentControlAnalayzer(
				new File("C:\\koda\\wordT\\samples\\SampleCreatedInWord1.docx"));
		wcca.scanDocument();
	}

}
