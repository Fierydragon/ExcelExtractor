package com.chen.li.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

public class ExcelUtilsTest {
	public static void main(String[] args) throws Exception {
		File f = new File("SampleSS.xlsx");
//	    Workbook wb = WorkbookFactory.create(f);
		InputStream in = new FileInputStream(f);
//		List<ExcelHeader2BeanFieldNameMapper> mappers = new ArrayList<ExcelHeader2BeanFieldNameMapper>();
		
		Pattern keyPattern = Pattern.compile("\\d+");
		Pattern valuePattern = Pattern.compile("\\w+");
		
		ExcelHeader2BeanFieldNameMapper keyMapper = new ExcelHeader2BeanFieldNameMapper("key", "key", keyPattern);
		ExcelHeader2BeanFieldNameMapper valueMapper = new ExcelHeader2BeanFieldNameMapper("value", "value", valuePattern);
		
		List<ExcelHeader2BeanFieldNameMapper> mappers = Arrays.asList(keyMapper, valueMapper);
		List<ErrorMessage> errorMessageList = new ArrayList<ErrorMessage>();
		List<TestBean> beanList = ExcelUtils.readExcelToEntity(TestBean.class, f, mappers, errorMessageList);
		System.out.println(beanList);
		System.out.println(errorMessageList);
	}
}
