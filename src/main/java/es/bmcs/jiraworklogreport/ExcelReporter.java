package es.bmcs.jiraworklogreport;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.joda.time.DateTime;

public class ExcelReporter {
	private List<String> fieldNames = new ArrayList<String>();
	private Workbook workbook = null;
	private String workbookName = "workbook.xls";

	public ExcelReporter(String workbookName) {
		setWorkbookName(workbookName);
		initialize();
	}

	private void initialize() {
		setWorkbook(new HSSFWorkbook());
	}

	public void closeWorksheet() {
		// Write the output to a file
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(getWorkbookName());
			getWorkbook().write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private boolean setupFieldsForClass(Class<?> clazz) throws Exception {
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			fieldNames.add(fields[i].getName());
		}
		return true;
	}

	private Sheet getSheetWithName(String name) {
		Sheet sheet = null;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			if (name.compareTo(workbook.getSheetName(i)) == 0) {
				sheet = workbook.getSheetAt(i);
				break;
			}
		}
		return sheet;
	}

	private void initializeForRead() throws InvalidFormatException, IOException {
		InputStream inp = new FileInputStream(getWorkbookName());
		workbook = WorkbookFactory.create(inp);
	}
	
	@SuppressWarnings("unchecked")
	public <T> List<T> readData(String classname) throws Exception {
		
		initializeForRead();
		Sheet sheet = getSheetWithName(classname);

		@SuppressWarnings("rawtypes")
		Class clazz = Class.forName(workbook.getSheetName(0));
		setupFieldsForClass(clazz);
		List<T> result = new ArrayList<T>();
		Row row;
		for (int rowCount = 1; rowCount < 4; rowCount++) {
			@SuppressWarnings("deprecation")
			T one = (T) clazz.newInstance();
			row = sheet.getRow(rowCount);
			int colCount = 0;
			result.add(one);
			for (Cell cell : row) {
				int type = cell.getCellType();
				String fieldName = fieldNames.get(colCount++);
				
				Method method = constructMethod(clazz, fieldName);
				
				if (type ==  1) {
					String value = cell.getStringCellValue();
					Object[] values = new Object[1];
					values[0] = value;
					method.invoke(one, values);
				} else if (type == 0) {
					Double num = cell.getNumericCellValue();
                    Class<?> returnType = getGetterReturnClass(clazz,fieldName);
                    if(returnType == Integer.class){
                    	method.invoke(one, num.intValue());
                    } else if(returnType == Double.class){
                    	method.invoke(one, num);
                    } else if(returnType == Float.class){
                    	method.invoke(one, num.floatValue());
                    }

				} else if (type == 3) {
					double num = cell.getNumericCellValue();
					Object[] values = new Object[1];
					values[0] =num;
					method.invoke(one, values);
				}
			}
		}

		return result;
	}

	private Class<?> getGetterReturnClass(Class<?> clazz, String fieldName) {
		String methodName = "get"+capitalize(fieldName);
		Class<?> returnType = null;
		for (Method method : clazz.getMethods()) {
			if(method.getName().equals(methodName)){
				returnType = method.getReturnType();
				break;
			}
		}
		return returnType;
	}
	@SuppressWarnings("unchecked")
	private Method constructMethod(Class clazz, String fieldName) throws SecurityException, NoSuchMethodException {
		Class<?> fieldClass = getGetterReturnClass(clazz, fieldName);
		return clazz.getMethod("set"+ capitalize(fieldName),fieldClass);
	}

	public <T> void writeReportToExcel(List<T> data) throws Exception {
		Sheet sheet = getWorkbook().createSheet(
				data.get(0).getClass().getName());
		setupFieldsForClass(data.get(0).getClass());
		// Create a row and put some cells in it. Rows are 0 based.
		int rowCount = 0;
		int columnCount = 0;

		Row row = sheet.createRow(rowCount++);
		for (String fieldName : fieldNames) {
			Cell cel = row.createCell(columnCount++);
			cel.setCellValue(fieldName);
		}
		Class<? extends Object> classz = data.get(0).getClass();
		for (T t : data) {
			row = sheet.createRow(rowCount++);
			columnCount = 0;
			for (String fieldName : fieldNames) {
				Cell cel = row.createCell(columnCount);
				Method method = classz.getMethod("get" + capitalize(fieldName));
				Object value = method.invoke(t, (Object[]) null);
				if (value != null) {
					if (value instanceof String) {
						cel.setCellValue((String) value);
					} else if (value instanceof Long) {
						cel.setCellValue((Long) value);
					} else if (value instanceof Integer) {
						cel.setCellValue((Integer)value);
					}else if (value instanceof Double) {
						cel.setCellValue((Double) value);
					} else if (value instanceof DateTime) {
						cel.setCellValue(new Date(((DateTime) value).getMillis()));
					}
				}
				columnCount++;
			}
		}
	}

	

	public String capitalize(String string) {
		String capital = string.substring(0, 1).toUpperCase();
		return capital + string.substring(1);
	}

	public String getWorkbookName() {
		return workbookName;
	}

	public void setWorkbookName(String workbookName) {
		this.workbookName = workbookName;
	}


	void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	Workbook getWorkbook() {
		return workbook;
	}

}