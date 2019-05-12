package com.chen.li.POI;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    private static final String FULL_DATA_FORMAT = "yyyy/MM/dd  HH:mm:ss";
    private static final String SHORT_DATA_FORMAT = "yyyy/MM/dd";
 
 
    /**
     * Excel表头对应Entity属性 解析封装javabean
     *
     * @param classzz    类
     * @param in         excel流
     * @param fileName   文件名
     * @param ExcelHeader2BeanFieldNameMappers excel表头与entity属性对应关系
     * @param <T>
     * @return
     * @throws Exception
     */
//    public static <T> List<T> readExcelToEntity(Class<T> classzz, InputStream in, String fileName, List<ExcelHeader2BeanFieldNameMapper> ExcelHeader2BeanFieldNameMappers) throws Exception {
//        checkFile(fileName);    //是否EXCEL文件
//        Workbook workbook = getWorkBoot(in, fileName); //兼容新老版本
//        List<T> excelForBeans = readExcel(classzz, workbook, ExcelHeader2BeanFieldNameMappers);  //解析Excel
//        return excelForBeans;
//    }
    

    
    /**
     * Excel表头对应Entity属性 解析封装javabean
     *
     * @param classzz    类
     * @param in         excel流
     * @param fileName   文件名
     * @param ExcelHeader2BeanFieldNameMappers excel表头与entity属性对应关系
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> readExcelToEntity(Class<T> classzz, File file, List<ExcelHeader2BeanFieldNameMapper> ExcelHeader2BeanFieldNameMappers, List<ErrorMessage> errorMessageList) throws Exception {
        checkFile(file.getName());    //是否EXCEL文件
//        Workbook workbook = getWorkBoot(file); //兼容新老版本
        Workbook workbook = WorkbookFactory.create(file);
        List<T> excelForBeans = readExcel(classzz, workbook, ExcelHeader2BeanFieldNameMappers, errorMessageList);  //解析Excel
        
        return excelForBeans;
    }
 
    /**
     * 解析Excel转换为Entity
     *
     * @param classzz  类
     * @param in       excel流
     * @param fileName 文件名
     * @param <T>
     * @return
     * @throws Exception
     */
//    public static <T> List<T> readExcelToEntity(Class<T> classzz, InputStream in, String fileName) throws Exception {
//        return readExcelToEntity(classzz, in, fileName,null);
//    }
 
    /**
     * 校验是否是Excel文件
     *
     * @param fileName
     * @throws Exception
     */
    public static void checkFile(String fileName) throws Exception {
        if (!StringUtils.isEmpty(fileName) && !(fileName.endsWith(".xlsx") || fileName.endsWith(".xls"))) {
            throw new Exception("不是Excel文件！");
        }
    }
 
    /**
     * 兼容新老版Excel
     *
     * @param in
     * @param fileName
     * @return
     * @throws IOException
     */
    private static Workbook getWorkBoot(InputStream in, String fileName) throws IOException {
        if (fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(in);
        } else {
            return new HSSFWorkbook(in);
        }
    }
    
//    /**
//     * 兼容新老版Excel
//     *
//     * @param in
//     * @param fileName
//     * @return
//     * @throws IOException
//     */
//    private static Workbook getWorkBoot(File file) throws IOException {
//        if (file.getName().endsWith(".xlsx")) {
//        	
//            return new XSSFWorkbook(in);
//        } else {
//            return new HSSFWorkbook(in);
//        }
//    }
 
    /**
     * 解析Excel
     *
     * @param classzz    类
     * @param workbook   工作簿对象
     * @param mappers excel与entity对应关系实体
     * @param <T>
     * @return
     * @throws Exception
     */
    private static <T> List<T> readExcel(Class<T> classzz, Workbook workbook, List<ExcelHeader2BeanFieldNameMapper> mappers, List<ErrorMessage> errorMessageList) throws Exception {
        List<T> beans = new ArrayList<T>();
        if (CollectionUtils.isEmpty(mappers)) {
            return null;
        }
        int sheetNum = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            String sheetName=sheet.getSheetName();
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            Row head = sheet.getRow(firstRowNum);
            if (head == null)
                continue;
            short firstCellNum = head.getFirstCellNum();
            short lastCellNum = head.getLastCellNum();
//            Field[] fields = classzz.getDeclaredFields();
            
            
            EgodicRows:
            for (int rowIndex = firstRowNum + 1; rowIndex <= lastRowNum; rowIndex++) {
                Row dataRow = sheet.getRow(rowIndex);
                if (dataRow == null)
                    continue;
                T instance = classzz.newInstance();
//                if(CollectionUtils.isEmpty(mappers)){  //非头部映射方式，默认不校验是否为空，提高效率
                    firstCellNum = dataRow.getFirstCellNum();
                    lastCellNum = dataRow.getLastCellNum();
//                }
//                for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
                
                if(dataRow.getLastCellNum() != mappers.size()) {
                	// TODO 记录错误信息
                	ErrorMessage errorMessage = new ErrorMessage(dataRow.getRowNum()+1, "The number of column and mapper are not corresponding!");
                	errorMessageList.add(errorMessage);
                	continue EgodicRows;
                }
                
                for(int index = 0; index < lastCellNum; index++) {
                    Cell cell = dataRow.getCell(index);
                    ExcelHeader2BeanFieldNameMapper mapper = mappers.get(index);
                    String cellValue = null;
                    if(cell == null) {
                    	if(mapper.isRequired()) {
                        	ErrorMessage errorMessage = new ErrorMessage(dataRow.getRowNum()+1, mapper.getExcelHeaderName() + "is required cell but now is null!");
                        	errorMessageList.add(errorMessage);
                        	continue EgodicRows;
                        } else {
                        	cellValue = "";
                        }
                    } else {
                    	cell.setCellType(CellType.STRING);
                        cellValue = cell.getStringCellValue();
                    }
                    
                	String fieldName = mapper.getBeanFieldName();
                	Field declaredField = classzz.getDeclaredField(fieldName);
                	String setMethodName = MethodUtils.setMethodName(fieldName);
                	Method setMethod = classzz.getMethod(setMethodName, declaredField.getType());
                	
                	if(isVolidated(cellValue, dataRow.getRowNum()+1, mapper, errorMessageList)) {
                		setMethod.invoke(instance, convertType(declaredField.getType(), cellValue.trim()));
                	} else {
                		continue EgodicRows;
                	}

                }
                beans.add(instance);
            }
        }
        return beans;
    }

	private static boolean isVolidated(String cellValue, int line, ExcelHeader2BeanFieldNameMapper mapper, List<ErrorMessage> errorMessageList) {
		// TODO Auto-generated method stub
		Matcher matcher = mapper.getValidPattern().matcher(cellValue);
		if(matcher.matches()) {
			return true;
		} else {
			ErrorMessage errorMessage = new ErrorMessage(line, mapper.getExcelHeaderName() + " Cell value is not valid!");
			errorMessageList.add(errorMessage);
			return false;
		}
		
	}


	/**
     * 是否日期字段
     *
     * @param field
     * @return
     */
    private static boolean isDateFied(Field field) {
        return (Date.class == field.getType());
    }
    /**
     * 空值校验
     *
     * @param ExcelHeader2BeanFieldNameMapper
     * @throws Exception
     */
//    private static void volidateValueRequired(ExcelHeader2BeanFieldNameMapper mapper,String sheetName,int rowIndex) throws Exception {
//        if (mapper != null && mapper.isRequired()) {
//            throw new Exception("《"+sheetName+"》第"+(rowIndex+1)+"行:\""+ mapper.getExcelHeaderName() + "\"不能为空！");
//        }
//    }
    /**
     * 类型转换
     *
     * @param classzz
     * @param value
     * @return
     */
    private static Object convertType(Class classzz, String value) {
        if (Integer.class == classzz || int.class == classzz) {
            return Integer.valueOf(value);
        }
        if (Short.class == classzz || short.class == classzz) {
            return Short.valueOf(value);
        }
        if (Byte.class == classzz || byte.class == classzz) {
            return Byte.valueOf(value);
        }
        if (Character.class == classzz || char.class == classzz) {
            return value.charAt(0);
        }
        if (Long.class == classzz || long.class == classzz) {
            return Long.valueOf(value);
        }
        if (Float.class == classzz || float.class == classzz) {
            return Float.valueOf(value);
        }
        if (Double.class == classzz || double.class == classzz) {
            return Double.valueOf(value);
        }
        if (Boolean.class == classzz || boolean.class == classzz) {
            return Boolean.valueOf(value.toLowerCase());
        }
        if (BigDecimal.class == classzz) {
            return new BigDecimal(value);
        }
       /* if (Date.class == classzz) {
            SimpleDateFormat formatter = new SimpleDateFormat(FULL_DATA_FORMAT);
            ParsePosition pos = new ParsePosition(0);
            Date date = formatter.parse(value, pos);
            return date;
        }*/
        return value;
    }
    /**
     * 获取properties的set和get方法
     */
    static class MethodUtils {
        private static final String SET_PREFIX = "set";
        private static final String GET_PREFIX = "get";
        private static String capitalize(String name) {
            if (name == null || name.length() == 0) {
                return name;
            }
            return name.substring(0, 1).toUpperCase() + name.substring(1);
        }
        public static String setMethodName(String propertyName) {
            return SET_PREFIX + capitalize(propertyName);
        }
        public static String getMethodName(String propertyName) {
            return GET_PREFIX + capitalize(propertyName);
        }
    }

}
