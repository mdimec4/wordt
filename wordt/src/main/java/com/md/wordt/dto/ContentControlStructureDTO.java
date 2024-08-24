package com.md.wordt.dto;

import java.util.ArrayList;
import java.util.List;

public class ContentControlStructureDTO {
	private String tag = "";
	private String name = "";
	private boolean repeating = false;
	private boolean multiParagraph = false;
	private List<ContentControlStructureDTO> children = new ArrayList<ContentControlStructureDTO>();

	public String getTag() {
		return tag;
	}

	public void setTag(String tag) {
		this.tag = tag;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public boolean isRepeating() {
		return repeating;
	}

	public void setRepeating(boolean repeating) {
		this.repeating = repeating;
	}

	public boolean isMultiParagraph() {
		return multiParagraph;
	}

	public void setMultiParagraph(boolean multiParagraph) {
		this.multiParagraph = multiParagraph;
	}

	public List<ContentControlStructureDTO> getChildren() {
		return children;
	}
}
