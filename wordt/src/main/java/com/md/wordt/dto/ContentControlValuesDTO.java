package com.md.wordt.dto;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ContentControlValuesDTO {
	private String label = "";
	private String value = "";
	
	private Map<String, List<ContentControlValuesDTO>> children = new HashMap<>();
	public String getLabel() {
		return label;
	}
	public void setLabel(String label) {
		this.label = label;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	public Map<String, List<ContentControlValuesDTO>> getChildren() {
		return children;
	}
	public void setChildren(Map<String, List<ContentControlValuesDTO>> children) {
		this.children = children;
	}
}