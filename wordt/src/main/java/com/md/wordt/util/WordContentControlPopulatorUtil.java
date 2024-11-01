package com.md.wordt.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTSdtCell;
import org.docx4j.wml.CTSdtRow;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtElement;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.tinylog.Logger;

import com.md.wordt.dto.util.ContentControlValuesDTO;

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

	private static P findFirstChildParagraphInContentAccessor(ContentAccessor contentAccessor) {
		Optional<Object> pOpt = contentAccessor.getContent() //
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
	private static Tc findFirstChildCellInSdtElement(SdtElement sdtElement) {
		Optional<Tc> tcOpt = sdtElement.getSdtContent().getContent() //
				.stream() //
				.filter(o -> {
					if (o instanceof Tc)
						return true;

					if (!(o instanceof JAXBElement))
						return false;

					JAXBElement el = (JAXBElement) o;
					if (!el.getDeclaredType().equals(Tc.class))
						return false;
					return true;
				}) //
				.map(o -> {
					if (o instanceof JAXBElement) {
						JAXBElement el = (JAXBElement) o;
						return (Tc) el.getValue();
					}
					return (Tc) o;
				}) //
				.findFirst();
		if (tcOpt.isEmpty())
			return null;
		return tcOpt.get();
	}

	private static R findFirstChildRunInParagrph(P p) {
		Optional<Object> rOpt = p.getContent() //
				.stream() //
				.filter(o -> (o instanceof R)) //
				.findFirst();
		if (rOpt.isEmpty())
			return null;
		return (R) rOpt.get();
	}

	@SuppressWarnings("rawtypes")
	private static Text findFirstTextInRun(R r) {
		Optional<Text> textOpt = r.getContent() //
				.stream() //
				.filter(o -> {
					if (o instanceof Text)
						return true;

					if (!(o instanceof JAXBElement))
						return false;
					JAXBElement rawJaxbEl = (JAXBElement) o;
					if (!rawJaxbEl.getDeclaredType().equals(Text.class))
						return false;
					return true;
				}) //
				.map(o -> {
					if (o instanceof Text)
						return (Text) o;
					JAXBElement rawJaxbEl = (JAXBElement) o;
					return (Text) rawJaxbEl.getValue();
				}) //
				.findFirst();
		if (textOpt.isEmpty())
			return null;

		return textOpt.get();
	}

	private static void populateWithText(SdtElement sdtElement, String textValue) {
		if (sdtElement == null) {
			Logger.error("populateWithText: null SdtElement");
			return;
		}
		if (textValue == null)
			textValue = "";

		if (sdtElement instanceof CTSdtCell) {
			populateWithTextSdtCell((CTSdtCell) sdtElement, textValue);
		} else if (sdtElement instanceof CTSdtRow) {
			Logger.error("populateWithText: CTSdtRow was not expected here!");
		} else if (sdtElement instanceof SdtBlock) {
			populateWithTextSdtBlock((SdtBlock) sdtElement, textValue);
		} else if (sdtElement instanceof SdtRun) {
			populateWithTextSdtRun((SdtRun) sdtElement, textValue);
		}
	}

	private static void populateWithTextSdtCell(CTSdtCell sdtCell, String textValue) {
		if (sdtCell == null) {
			Logger.error("populateWithTextSdtCell: null TDSdtCell");
			return;
		}

		if (textValue == null) {
			textValue = "";
		}
		Tc tc = findFirstChildCellInSdtElement(sdtCell);
		if (tc == null) {
			Logger.error("populateWithTextSdtCell: null Tc");
			return;
		}

		P p = findFirstChildParagraphInContentAccessor(tc);
		if (p == null) {
			Logger.error("populateWithTextSdtCell: null P");
			return;
		}

		R firstRun = findFirstChildRunInParagrph(p);
		if (firstRun == null) {
			Logger.error("populateWithTextSdtCell: null R");
			return;
		}

		populateWithTextRun(firstRun, textValue);
	}

	private static void populateWithTextSdtBlock(SdtBlock sdtBlock, String textValue) {
		if (sdtBlock == null) {
			Logger.error("populateWithTextSdtBlock: null SdtBlock");
			return;
		}

		if (textValue == null) {
			textValue = "";
		}
		textValue = textValue.replace("\r\n", "\n").replace("\r", "\n");
		String[] lines = textValue.split("\n"); // empty line means new paragraph
		P referenceP = findFirstChildParagraphInSdtElement(sdtBlock);
		if (referenceP == null) {
			Logger.error("populateWithTextSdtBlock: SdtBlock paragraph is missing!");
			return;
		}
		R referenceR = findFirstChildRunInParagrph(referenceP);

		sdtBlock.getSdtContent().getContent().clear();

		P currentP = null;
		for (String line : lines) {
			if (currentP == null || line.isEmpty()) {
				currentP = createP(sdtBlock.getSdtContent(), referenceP);
				sdtBlock.getSdtContent().getContent().add(currentP);
			}

			boolean lineBreak = false;
			if (!currentP.getContent().isEmpty() && !line.isEmpty()) {
				R newLineR = createNewLineR(currentP);
				currentP.getContent().add(newLineR);
				lineBreak = true;
			}

			if ((!lineBreak && line.isEmpty()) || !line.isEmpty()) {
				R rWithText = createR(currentP, referenceR, line);
				currentP.getContent().add(rWithText);
			}

			if (!lineBreak && line.isEmpty())
				currentP = null;

		}
	}

	private static P createP(Object parent, P referenceP) {
		P p = new P();
		if (referenceP != null && referenceP.getPPr() != null) {
			PPr ppr = XmlUtils.deepCopy(referenceP.getPPr());
			p.setPPr(ppr);
		}
		p.setParent(parent);
		return p;
	}

	private static R createR(Object parent, R referenceR, String contentText) {
		if (contentText == null)
			contentText = "";
		R r = new R();
		if (referenceR != null && referenceR.getRPr() != null) {
			RPr rpr = XmlUtils.deepCopy(referenceR.getRPr());
			rpr.setRStyle(null);

			r.setRPr(rpr);
		}
		r.setParent(parent);

		Text text = new Text();
		text.setParent(r);
		r.getContent().add(text);

		text.setValue(contentText);
		return r;
	}

	private static R createNewLineR(Object parent) {
		R r = new R();
		r.setParent(parent);
		Br br = new Br();
		br.setParent(r);
		r.getContent().add(br);
		return r;
	}

	private static void populateWithTextSdtRun(SdtRun sdtRun, String textValue) {
		if (sdtRun == null) {
			Logger.error("populateWithTextSdtRun: null SdtRun");
			return;
		}

		if (textValue == null) {
			textValue = "";
		}

		R firstRun = findFirstChildRunInSdtElement(sdtRun);
		if (firstRun == null) {
			Logger.error("populateWithTextSdtRun: null R");
			return;
		}

		sdtRun.getSdtContent().getContent().clear();
		sdtRun.getSdtContent().getContent().add(firstRun); // remove all other content

		populateWithTextRun(firstRun, textValue);
	}

	private static void populateWithTextRun(R r, String textValue) {
		if (r == null) {
			Logger.error("populateWithTextRun: null R");
			return;
		}

		if (textValue == null) {
			textValue = "";
		}
		r.getRPr().setRStyle(null);
		Text firstTextInRun = findFirstTextInRun(r);
		if (firstTextInRun == null) {
			Logger.error("populateWithTextRun: child Text element not found");
			return;
		}

		textValue = textValue.replace("\r", "").replace("\n", "");

		firstTextInRun.setValue(textValue);
	}

	private static int getNumberOfRepeatedItemsFromDTO(ContentControlValuesDTO repeatingDTO) {
		if (repeatingDTO == null || repeatingDTO.getChildren() == null || repeatingDTO.getChildren().keySet().isEmpty())
			return 0;

		int minNumber = Integer.MAX_VALUE;
		for (String key : repeatingDTO.getChildren().keySet()) {
			List<ContentControlValuesDTO> children = repeatingDTO.getChildren().get(key);
			int num = children.size();

			if (num < minNumber)
				minNumber = num;
		}
		return minNumber;
	}

	private static void populateDocumentRecursive(List<Object> docElementContent,
			Map<String, List<ContentControlValuesDTO>> dtoContent, boolean parenSdtIsRepeatingSectionItem,
			Map<String, Integer> repeatingSectionitemIndexMap, String repitingSectionTag) {
		if (docElementContent == null || dtoContent == null)
			return;
		Map<String, Integer> repeatingSectionitemIndexMapForNextGeneration = repeatingSectionitemIndexMap;

		for (int i = 0; i < docElementContent.size(); i++) {
			Object object = docElementContent.get(i);
			if (object instanceof JAXBElement<?>) {
				object = ((JAXBElement<?>) object).getValue();
			}
			// System.out.println(object.getClass().getSimpleName());

			boolean forNextGenerationIsSdtRepeatingSectionItem = parenSdtIsRepeatingSectionItem;
			String nextRepeatingSectionTag = repitingSectionTag;

			List<Object> childContent = null;
			Map<String, List<ContentControlValuesDTO>> childDTOs = null;

			if (object instanceof SdtElement) {
				SdtElement sdtElement = (SdtElement) object;
				String tag = WordContentControlAnalayzerUtil.getSdtTag(sdtElement);

				int repeatingSectionitemIndex = 0;
				if (!tag.isEmpty()) {
					if (repeatingSectionitemIndexMap != null) {
						if (repeatingSectionitemIndexMap.get(tag) != null)
							repeatingSectionitemIndex = repeatingSectionitemIndexMap.get(tag);
						else
							repeatingSectionitemIndexMap.put(tag, repeatingSectionitemIndex);
					}
				}

				childContent = sdtElement.getSdtContent().getContent();

				// System.out.println("Tag: " + tag);
				// String name = WordContentControlAnalayzerUtil.getSdtName(sdtElement);
				// System.out.println("Name: " + name);
				boolean isRepeatingSection = WordContentControlAnalayzerUtil.isW15RepeatingSection(sdtElement);
				if (isRepeatingSection) {
					repeatingSectionitemIndexMapForNextGeneration = new HashMap<String, Integer>();
					nextRepeatingSectionTag = tag;
				}
				// System.out.println("isRepeatingSection: " + isRepeatingSection);
				boolean isRepeatingSectionItem = WordContentControlAnalayzerUtil.isW15RepeatingSectionItem(sdtElement);

				forNextGenerationIsSdtRepeatingSectionItem = isRepeatingSectionItem;

				// System.out.println("isRepatingSectionItem: " + isRepeatingSectionItem);
				// boolean isMultiParagraph =
				// WordContentControlAnalayzerUtil.isMultipleParagraph(sdtElement);
				// System.out.println("MultiParagraph: " + isMultiParagraph);

				if (!tag.isEmpty()) {
					List<ContentControlValuesDTO> dtos = dtoContent.get(tag);
					if (dtos == null || dtos.isEmpty())
						continue;

					if (isRepeatingSection) {
						int expectedNumber = getNumberOfRepeatedItemsFromDTO(dtos.get(0));
						if (expectedNumber > 0) { //we have some children
							int haveNumber = childContent.size();
							int toAddNumber = expectedNumber - haveNumber;
							if (haveNumber > 0) {
								SdtElement repeatItem = WordContentControlAnalayzerUtil
										.findRepeatingSectionItemChild(sdtElement);
	
								if (repeatItem != null) {
									for (int j = 0; j < toAddNumber; j++) {
										SdtElement repeatItemCopy = XmlUtils.deepCopy(repeatItem);
										childContent.add(repeatItemCopy);
										// System.out.println("Add copy");
									}
								}
							}
						} else {
							int noRepeats = dtos.size();
							if (noRepeats > 0) { //we have some children
								int haveNumber = childContent.size();
								int toAddNumber = noRepeats - haveNumber;
								if (haveNumber > 0) {
									SdtElement repeatItem = WordContentControlAnalayzerUtil
											.findRepeatingSectionItemChild(sdtElement);
		
									if (repeatItem != null) {
										for (int j = 0; j < toAddNumber; j++) {
											SdtElement repeatItemCopy = XmlUtils.deepCopy(repeatItem);
											childContent.add(repeatItemCopy);
											// System.out.println("Add copy");
										}
									}
								}
							}
						}

					} else if (dtos.size() >= (repeatingSectionitemIndex + 1)) {

						String textValue = dtos.get(repeatingSectionitemIndex).getValue();
						if (textValue == null)
							textValue = "";
						populateWithText(sdtElement, textValue);
					}

					if (dtos.size() >= (repeatingSectionitemIndex + 1) && parenSdtIsRepeatingSectionItem) {
						childDTOs = dtos.get(repeatingSectionitemIndex).getChildren();
					}
				} else if (isRepeatingSectionItem && repitingSectionTag != null && !repitingSectionTag.isEmpty()) {
					List<ContentControlValuesDTO> dtos = dtoContent.get(repitingSectionTag);
					if (dtos == null || dtos.isEmpty())
						continue;
					if (dtos.size() >= (i + 1)) {
						String textValue = dtos.get(i).getValue();
						if (textValue == null)
							textValue = "";
						populateWithText(sdtElement, textValue);
					}
				}

				if (parenSdtIsRepeatingSectionItem) {
					repeatingSectionitemIndex++;
					if (!tag.isEmpty()) {
						if (repeatingSectionitemIndexMapForNextGeneration != null) {
							repeatingSectionitemIndexMapForNextGeneration.put(tag, repeatingSectionitemIndex);
						}
					}

				}

			} else if (object instanceof ContentAccessor) {
				ContentAccessor ca = (ContentAccessor) object;
				childContent = ca.getContent();
			}

			if (childContent != null && childDTOs == null)
				populateDocumentRecursive(childContent, dtoContent, forNextGenerationIsSdtRepeatingSectionItem,
						repeatingSectionitemIndexMapForNextGeneration, nextRepeatingSectionTag);
			else if (childContent != null && childDTOs != null)
				populateDocumentRecursive(childContent, childDTOs, forNextGenerationIsSdtRepeatingSectionItem,
						repeatingSectionitemIndexMapForNextGeneration, nextRepeatingSectionTag);
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
			populateDocumentRecursive(((ContentAccessor) headerPart.getContents()).getContent(), dtoContent, false,
					null, null);
		}
		// get main part placeholders
		populateDocumentRecursive(mainDocumentPart.getContent(), dtoContent, false, null, null);

		// get footer placeholders
		for (JaxbXmlPart footerPart : footerParts) {
			populateDocumentRecursive(((ContentAccessor) footerPart.getContents()).getContent(), dtoContent, false,
					null, null);
		}
	}
}
