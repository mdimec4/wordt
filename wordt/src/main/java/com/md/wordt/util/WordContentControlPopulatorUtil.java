package com.md.wordt.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

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
	public void saveAsPdf(OutputStream outputStream) throws Docx4JException {
		Docx4J.toPDF(wordPackage, outputStream);
	}

	public void saveAsPdf(File file) throws Docx4JException, FileNotFoundException {
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		saveAsPdf(fileOutputStream);
	}
	 */
}
