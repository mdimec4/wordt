package com.md.wordt.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTSdtCell;
import org.docx4j.wml.CTSdtRow;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtElement;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Text;
import org.tinylog.Logger;

import com.md.wordt.dto.ContentControlValuesDTO;

import jakarta.xml.bind.JAXBElement;

public class WordContentControlPopulatorUtil {

	private WordprocessingMLPackage wordPackage;

	public WordContentControlPopulatorUtil(WordprocessingMLPackage wordPackage) {
		this.wordPackage = wordPackage;

	}

	public WordContentControlPopulatorUtil(InputStream inputStream) throws Docx4JException {
		wordPackage = Docx4J.load(inputStream);
	}

	public WordContentControlPopulatorUtil(File file) throws Docx4JException, FileNotFoundException {
		this(new FileInputStream(file));
	}

	public void save(OutputStream outputStream) throws Docx4JException {
		wordPackage.save(outputStream);
	}

	public void save(File file) throws Docx4JException {
		wordPackage.save(file);
	}

	/*
	 * public void saveAsPdf(OutputStream outputStream) throws Docx4JException {
	 * Docx4J.toPDF(wordPackage, outputStream); }
	 * 
	 * public void saveAsPdf(File file) throws Docx4JException,
	 * FileNotFoundException { FileOutputStream fileOutputStream = new
	 * FileOutputStream(file); saveAsPdf(fileOutputStream); }
	 */

	private static P findFirstChildParagraphInSdtElement(SdtElement sdtElement) {
		Optional<Object> pOpt = sdtElement.getSdtContent().getContent() //
				.stream() //
				.filter(o -> (o instanceof P)) //
				.findFirst();
		if (pOpt.isEmpty())
			return null;
		return (P) pOpt.get();
	}

	private static R findFirstChildRunInSdtElement(SdtElement sdtElement) {
		Optional<Object> rOpt = sdtElement.getSdtContent().getContent() //
				.stream() //
				.filter(o -> (o instanceof R)) //
				.findFirst();
		if (rOpt.isEmpty())
			return null;
		return (R) rOpt.get();
	}

	@SuppressWarnings("rawtypes")
	private static Text findFirstTextInRun(R r) {
		Optional<Object> jaxbElmOpt = r.getContent() //
				.stream() //
				.filter(o -> (o instanceof JAXBElement)).findFirst();
		if (jaxbElmOpt.isEmpty())
			return null;
		JAXBElement jaxbElem = (JAXBElement) jaxbElmOpt.get();
		if (!jaxbElem.getDeclaredType().equals(Text.class))
			return null;
		return (Text) jaxbElem.getValue();
	}

	private static void populateWithText(SdtElement sdtElement, String textValue) {
		if (sdtElement == null || textValue == null || textValue.isEmpty())
			return;

		if (sdtElement instanceof CTSdtCell) {
		} else if (sdtElement instanceof CTSdtRow) {
		} else if (sdtElement instanceof SdtBlock) {
		} else if (sdtElement instanceof SdtRun) {
			populateWithTextSdtRun((SdtRun) sdtElement, textValue);
		}
	}

	private static void populateWithTextSdtRun(SdtRun sdtRun, String textValue) {
		if (sdtRun == null || textValue == null || textValue.isEmpty())
			return;

		R firstRun = findFirstChildRunInSdtElement(sdtRun);

		sdtRun.getSdtContent().getContent().clear();
		sdtRun.getSdtContent().getContent().add(firstRun); // remove all other content

		populateWithTextRun(firstRun, textValue);
	}

	private static void populateWithTextRun(R r, String textValue) {
		r.getRPr().setRStyle(null);
		Text firstTextInRun = findFirstTextInRun(r);
		if (firstTextInRun == null)
			return;

		textValue = textValue.replace("\r", "").replace("\n", "");

		firstTextInRun.setValue(textValue);
	}

	private static void populateDocumentRecursive(List<Object> docElementContent,
			Map<String, List<ContentControlValuesDTO>> dtoContent) {
		if (docElementContent == null || dtoContent == null)
			return;

		for (Object object : docElementContent) {
			if (object instanceof JAXBElement<?>) {
				object = ((JAXBElement<?>) object).getValue();
			}
			System.out.println(object.getClass().getSimpleName());

			List<Object> childContent = null;
			Map<String, List<ContentControlValuesDTO>> childDTOs = null;
			if (object instanceof SdtElement) {
				SdtElement sdtElement = (SdtElement) object;
				String tag = WordContentControlAnalayzerUtil.getSdtTag(sdtElement);

				childContent = sdtElement.getSdtContent().getContent();

				// System.out.println("Tag: " + tag);
				String name = WordContentControlAnalayzerUtil.getSdtName(sdtElement);
				// System.out.println("Name: " + name);
				boolean isRepeatingSection = WordContentControlAnalayzerUtil.isW15RepeatingSection(sdtElement);
				// System.out.println("isRepeatingSection: " + isRepeatingSection);
				boolean isRepeatingSectionItem = WordContentControlAnalayzerUtil.isW15RepeatingSectionItem(sdtElement);
				// System.out.println("isRepatingSectionItem: " + isRepeatingSectionItem);
				boolean isMultiParagraph = WordContentControlAnalayzerUtil.isMultipleParagraph(sdtElement);
				System.out.println("MultiParagraph: " + isMultiParagraph);

				if (!tag.isEmpty()) {
					List<ContentControlValuesDTO> dtos = dtoContent.get(tag);
					if (dtos == null || dtos.isEmpty())
						continue;

					if (isRepeatingSection) {

					} else {
						if (dtos.size() > 1) {
							Logger.error(
									"For label {} there are multiple entries for non repeating section. Extra will be ignored!",
									tag);
						}
						String textValue = dtos.get(0).getValue();
						if (textValue != null && !textValue.isEmpty())
							populateWithText(sdtElement, textValue);
					}

					// childDTOs = dtoContent.get(tag).getChildren();
				}

			} else if (object instanceof ContentAccessor) {
				ContentAccessor ca = (ContentAccessor) object;
				childContent = ca.getContent();
			}

			if (childContent != null && childDTOs == null)
				populateDocumentRecursive(childContent, dtoContent);
			else if (childContent != null && childDTOs != null)
				populateDocumentRecursive(childContent, childDTOs);
		}
	}

	@SuppressWarnings("rawtypes")
	public void populateDocument(Map<String, List<ContentControlValuesDTO>> dtoContent) throws Docx4JException {

		MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
		List<JaxbXmlPart> headerParts = new ArrayList<JaxbXmlPart>();
		List<JaxbXmlPart> footerParts = new ArrayList<JaxbXmlPart>();

		// get also header placeholders
		RelationshipsPart relationshipPart = mainDocumentPart.getRelationshipsPart();
		List<Relationship> relationships = relationshipPart.getRelationships().getRelationship();
		for (Relationship r : relationships) {
			if (r.getType().equals(Namespaces.HEADER)) {
				headerParts.add((JaxbXmlPart) relationshipPart.getPart(r));
			} else if (r.getType().equals(Namespaces.FOOTER)) {
				footerParts.add((JaxbXmlPart) relationshipPart.getPart(r));
			}
		}

		for (JaxbXmlPart headerPart : headerParts) {
			populateDocumentRecursive(((ContentAccessor) headerPart.getContents()).getContent(), dtoContent);
		}
		// get main part placeholders
		populateDocumentRecursive(mainDocumentPart.getContent(), dtoContent);

		// get footer placeholders
		for (JaxbXmlPart footerPart : footerParts) {
			populateDocumentRecursive(((ContentAccessor) footerPart.getContents()).getContent(), dtoContent);
		}
	}
}
