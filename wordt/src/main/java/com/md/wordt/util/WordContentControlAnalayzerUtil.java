package com.md.wordt.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

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
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtElement;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tag;

import com.md.wordt.dto.ContentControlStructureDTO;
import com.md.wordt.dto.ContentControlValuesDTO;

import jakarta.xml.bind.JAXBElement;

public class WordContentControlAnalayzerUtil {

	private WordprocessingMLPackage wordPackage;

	public WordContentControlAnalayzerUtil(WordprocessingMLPackage wordPackage) {
		this.wordPackage = wordPackage;
	}

	public WordContentControlAnalayzerUtil(InputStream inputStream) throws Docx4JException {
		wordPackage = Docx4J.load(inputStream);
	}

	public WordContentControlAnalayzerUtil(File file) throws Docx4JException, FileNotFoundException {
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

	public static boolean isMultipleParagraph(SdtElement sdtElement) {
		if (sdtElement instanceof CTSdtCell) {
			return false;
		} else if (sdtElement instanceof CTSdtRow) {
			return false;
		} else if (sdtElement instanceof SdtBlock) {
			return true;
		} else if (sdtElement instanceof SdtRun) {
			return false;
		}

		return false;
	}

	private static void scanDocumentRecursive(List<Object> docElementContent,
			List<ContentControlStructureDTO> dtoContent, Set<String> levelContentControlDuplicatePreventingSet) {
		if (docElementContent == null || dtoContent == null || levelContentControlDuplicatePreventingSet == null)
			return;

		for (Object object : docElementContent) {
			if (object instanceof JAXBElement<?>) {
				object = ((JAXBElement<?>) object).getValue();
			}
			// System.out.println(object.getClass().getSimpleName());

			List<Object> childContent = null;
			List<ContentControlStructureDTO> childDTOs = null;
			if (object instanceof SdtElement) {
				SdtElement sdtElement = (SdtElement) object;
				String tag = getSdtTag(sdtElement);

				if (!tag.isEmpty()) {
					if (levelContentControlDuplicatePreventingSet.contains(tag))
						continue;
					levelContentControlDuplicatePreventingSet.add(tag);
				}

				childContent = sdtElement.getSdtContent().getContent();

				// System.out.println("Tag: " + tag);
				String name = getSdtName(sdtElement);
				// System.out.println("Name: " + name);
				boolean isRepeatingSection = isW15RepeatingSection(sdtElement);
				// System.out.println("isRepeatingSection: " + isRepeatingSection);
				// boolean isRepeatingSectionItem = isW15RepeatingSectionItem(sdtElement);
				// System.out.println("isRepatingSectionItem: " + isRepeatingSectionItem);

				boolean isMultiParagraph = isMultipleParagraph(sdtElement);

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
				scanDocumentRecursive(childContent, dtoContent, levelContentControlDuplicatePreventingSet);
			else if (childContent != null && childDTOs != null)
				scanDocumentRecursive(childContent, childDTOs, new HashSet<String>());
		}
	}

	@SuppressWarnings("rawtypes")
	public List<ContentControlStructureDTO> scanDocument() throws Docx4JException {
		List<ContentControlStructureDTO> rootList = new ArrayList<ContentControlStructureDTO>();
		Set<String> rootLevelContentControlDuplicatePreventingSet = new HashSet<String>();

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
			scanDocumentRecursive(((ContentAccessor) headerPart.getContents()).getContent(), rootList,
					rootLevelContentControlDuplicatePreventingSet);
		}
		// get main part placeholders
		scanDocumentRecursive(mainDocumentPart.getContent(), rootList, rootLevelContentControlDuplicatePreventingSet);

		// get footer placeholders
		for (JaxbXmlPart footerPart : footerParts) {
			scanDocumentRecursive(((ContentAccessor) footerPart.getContents()).getContent(), rootList,
					rootLevelContentControlDuplicatePreventingSet);
		}

		return rootList;
	}

	public static void main(String[] args) throws FileNotFoundException, Docx4JException {
		WordprocessingMLPackage wordPackageSource = Docx4J
				.load(new File("C:\\koda\\wordT\\samples\\SampleCreatedInWord1.docx"));

		WordContentControlAnalayzerUtil wcca = new WordContentControlAnalayzerUtil(wordPackageSource);
		List<ContentControlStructureDTO> structureDTOs = wcca.scanDocument();

		// clone
		WordprocessingMLPackage wordPackageTarget = (WordprocessingMLPackage) wordPackageSource.clone();

		Map<String, List<ContentControlValuesDTO>> toPupulateDTOs = new HashMap<>();

		ContentControlValuesDTO headerCc = new ContentControlValuesDTO();
		headerCc.setLabel("HEADER");
		headerCc.setValue("Muc je header");
		toPupulateDTOs.put(headerCc.getLabel(), List.of(headerCc));

		ContentControlValuesDTO runCc = new ContentControlValuesDTO();
		runCc.setLabel("RUN");
		runCc.setValue("Niko je muc!");
		toPupulateDTOs.put(runCc.getLabel(), List.of(runCc));

		ContentControlValuesDTO richCc = new ContentControlValuesDTO();
		richCc.setLabel("RICH_ONE");
		richCc.setValue("\n" //
				+ "Runo mali\n" //
				+ "pesek moj zali\n" //
				+ "\n" //
				+ "na travniky sedi\n" //
				+ "in se veseli!");
		toPupulateDTOs.put(richCc.getLabel(), List.of(richCc));

		ContentControlValuesDTO tabela1Cc = new ContentControlValuesDTO();
		tabela1Cc.setLabel("TABELA1"); // TODO ne dela
		tabela1Cc.setValue("Muc je muc");
		toPupulateDTOs.put(tabela1Cc.getLabel(), List.of(tabela1Cc));

		ContentControlValuesDTO tabela2Cc = new ContentControlValuesDTO();
		tabela2Cc.setLabel("TABELA2"); // TODO ne dela
		tabela2Cc.setValue("Muc je muc");
		toPupulateDTOs.put(tabela2Cc.getLabel(), List.of(tabela2Cc));

		ContentControlValuesDTO footerCc = new ContentControlValuesDTO();
		footerCc.setLabel("FOOTER");
		footerCc.setValue("Muc je footer");
		toPupulateDTOs.put(footerCc.getLabel(), List.of(footerCc));

		WordContentControlPopulatorUtil populator = new WordContentControlPopulatorUtil(wordPackageTarget);
		populator.populateDocument(toPupulateDTOs);

		populator.save(new File("C:\\koda\\wordT\\samples\\SampleCreatedInWord1_NIKO.docx"));
	}

}
