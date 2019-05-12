package com.chen.li.POI;

import java.util.regex.Pattern;

public class ExcelHeader2BeanFieldNameMapper {
	private String excelHeaderName;
	private String beanFieldName;
	private Pattern validPattern;
	public boolean required = true;
	
	public ExcelHeader2BeanFieldNameMapper(String excelHeaderName, String beanFieldName) {
		this(excelHeaderName, beanFieldName, null, false);
	}
	
	public ExcelHeader2BeanFieldNameMapper(String excelHeaderName, String beanFieldName, boolean requied) {
		this(excelHeaderName, beanFieldName, null, requied);
	}
	
	public ExcelHeader2BeanFieldNameMapper(String excelHeaderName, String beanFieldName, Pattern validPattern) {
		this(excelHeaderName, beanFieldName, validPattern, true);
	}
	
	public ExcelHeader2BeanFieldNameMapper(String excelHeaderName, String beanFieldName, Pattern validPattern,
			boolean requied) {
		this.excelHeaderName = excelHeaderName;
		this.beanFieldName = beanFieldName;
		this.validPattern = validPattern;
		this.required = requied;
	}

	public String getExcelHeaderName() {
		return excelHeaderName;
	}
	public void setExcelHeaderName(String excelHeaderName) {
		this.excelHeaderName = excelHeaderName;
	}
	public String getBeanFieldName() {
		return beanFieldName;
	}
	public void setBeanFieldName(String beanFieldName) {
		this.beanFieldName = beanFieldName;
	}
	public Pattern getValidPattern() {
		return validPattern;
	}
	public void setValidPattern(Pattern validPattern) {
		this.validPattern = validPattern;
	}
	public boolean isRequired() {
		return required;
	}
	public void setRequired(boolean requied) {
		this.required = requied;
	}
}
