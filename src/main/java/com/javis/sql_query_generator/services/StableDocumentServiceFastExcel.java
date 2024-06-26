package com.javis.sql_query_generator.services;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.*;

public class StableDocumentServiceFastExcel {
	public static StableDocumentXls readFromWorkBook(ReadableWorkbook workbook) throws IOException, ParseException {
		return readFromWorkBook(workbook, null);
	}

	/**
	 * read only till {@code maxRows} */
	public static StableDocumentXls readFromWorkBook(ReadableWorkbook workbook, Integer maxRows) throws IOException, ParseException {
		List<String> headers = new ArrayList<>();
		List<List<String>> data = new ArrayList<>();
		Map<String, Integer> columnIndexMap = new HashMap<>();
		Map<Integer, String> indexColumnMap = new HashMap<>();
		Map<String, Integer> headerNameMap = new HashMap<>();
		Sheet firstSheet = workbook.getFirstSheet();
		int rowCount = 0;
		int headerCount = 0;
		for (Row row : firstSheet.read()) {
			List<String> rowData = new ArrayList<>();
			headerCount = Math.max(row.getCellCount(),headerCount);
//            headerCount = headers.size()==0?rowData.size():headers.size();
//			System.out.println(rowCount);
			for (int columnIndex = 0; columnIndex < headerCount; columnIndex++) {

				if (row.getCellCount() > columnIndex) {
//				System.out.println("h "+headerCount);
					Cell cell = row.getCell(columnIndex);

					if (rowCount == 0) {
						String value = null;
						if (cell != null && cell.getValue() != null) {
							value = cell.getValue().toString().trim();
							headers.add(value);
						}

						headerNameMap.put(value, columnIndex);
						columnIndexMap.put(value,columnIndex);
						indexColumnMap.put(columnIndex, value);

					} else {

						if (cell != null && cell.getValue() != null) {
							if(cell.getType().toString().equals("NUMBER")){
								Double cellValue = Double.valueOf(cell.getValue().toString());
								if(cellValue % 1 == 0){
									Integer intVal = cellValue.intValue();
									rowData.add(intVal.toString());
								}
								else{
									rowData.add(cellValue.toString());
								}
							}
							else{
								rowData.add(cell.getValue().toString());
							}

						} else {
							rowData.add("");
						}
					}
				}
				else{
					rowData.add("");
				}
			}

			if (rowCount != 0 && !checkEmptyRow(rowData)) {
				data.add(rowData);
			}

			rowCount++;
			if (maxRows != null && rowCount == maxRows)
				break;
		}

		StableDocumentXls xls = StableDocumentXls.builder()
				.headers(headers)
				.indexColumnMap(indexColumnMap)
				.columnIndexMap(columnIndexMap)
				.data(data)
				.build();
		return xls;
	}

	public static StableDocumentXls readFromWorkBook(ReadableWorkbook workbook, Integer maxRows, Sheet sheet) throws IOException, ParseException {
		List<String> headers = new ArrayList<>();
		List<List<String>> data = new ArrayList<>();
		Map<String, Integer> columnIndexMap = new HashMap<>();
		Map<Integer, String> indexColumnMap = new HashMap<>();
		Map<String, Integer> headerNameMap = new HashMap<>();
		int rowCount = 0;
		int headerCount = 0;
		for (Row row : sheet.read()) {
			List<String> rowData = new ArrayList<>();
			headerCount = Math.max(row.getCellCount(),headerCount);
//            headerCount = headers.size()==0?rowData.size():headers.size();
//			System.out.println(rowCount);
			for (int columnIndex = 0; columnIndex < headerCount; columnIndex++) {

				if (row.getCellCount() > columnIndex) {
//				System.out.println("h "+headerCount);
					Cell cell = row.getCell(columnIndex);

					if (rowCount == 0) {
						String value = null;
						if (cell != null && cell.getValue() != null) {
							value = cell.getValue().toString().trim();
							headers.add(value);
						}

						headerNameMap.put(value, columnIndex);
						columnIndexMap.put(value,columnIndex);
						indexColumnMap.put(columnIndex, value);

					} else {

						if (cell != null && cell.getValue() != null) {
							if(cell.getType().toString().equals("NUMBER")){
								Double cellValue = Double.valueOf(cell.getValue().toString());
								if(cellValue % 1 == 0){
									Integer intVal = cellValue.intValue();
									rowData.add(intVal.toString());
								}
								else{
									rowData.add(cellValue.toString());
								}
							}
							else{
								rowData.add(cell.getValue().toString());
							}

						} else {
							rowData.add("");
						}
					}
				}
				else{
					rowData.add("");
				}
			}

			if (rowCount != 0 && !checkEmptyRow(rowData)) {
				data.add(rowData);
			}

			rowCount++;
			if (maxRows != null && rowCount == maxRows)
				break;
		}

		StableDocumentXls xls = StableDocumentXls.builder()
				.headers(headers)
				.indexColumnMap(indexColumnMap)
				.columnIndexMap(columnIndexMap)
				.data(data)
				.build();
		return xls;
	}



	private static boolean checkEmptyRow(List<String> rowData) {
		for (String data : rowData) {
			if (!data.trim().isEmpty()) {
				return false;
			}
		}
		return true;
	}

	// Create workbook from Document object
	public static void setColumnMap(StableDocumentXls documentXls){
		Map<String, Integer> columnIndexMap = new HashMap<>();
		Map<Integer, String> indexColumnMap = new HashMap<>();
		int position = 0;
		for(String header: documentXls.getHeaders()){
			columnIndexMap.put(header, position);
			indexColumnMap.put(position, header);
			position++;
		}
		documentXls.setColumnIndexMap(columnIndexMap);
		documentXls.setIndexColumnMap(indexColumnMap);
	}



	private static void insertRow(Worksheet firstSheet, List<String> data, int rowNum) {
		int cellnum = 0;
		for (String obj : data) {

			if(Objects.equals(obj, "TRUE()")||Objects.equals(obj, "true()")||Objects.equals(obj, "True()") )
				firstSheet.value(rowNum, cellnum, true);
			else if(Objects.equals(obj,"FALSE()")||Objects.equals(obj, "false()")||Objects.equals(obj, "False()"))
				firstSheet.value(rowNum, cellnum, false);
			else
				firstSheet.value(rowNum, cellnum, obj);
			cellnum++;

		}
	}

	public static void populateDatesColumns(Set<String> s){
		s.add("Date");
		s.add("Valid From");
		s.add("Valid To");
	}

	public static Boolean isDateColumn(String columnName, Set<String> dateColumns){
		return dateColumns.contains(columnName)||columnName.endsWith("Date")||columnName.endsWith("date");
	}
	// Converting File to Workbook

	public static Workbook writeToWorkBook(StableDocumentXls xls) throws IOException {

		var f = new File("alternate-csv-18.xlsx");

		var fos = new FileOutputStream(f);
		Workbook workbook = new Workbook(fos, "workbook","1.0");
		Worksheet firstSheet = workbook.newWorksheet("Sheet1");
		insertRow(firstSheet, xls.getHeaders(), 0);
		int rowNum = 1;
		if (xls.getData().size() > 0) {
			for (List<String> rowData : xls.getData()) {
				insertRow(firstSheet, rowData, rowNum);
				rowNum++;
			}
		}
		workbook.finish();
		return workbook;
	}

	public static ByteArrayOutputStream writeToWorkBook1(StableDocumentXls xls) throws IOException {

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		Workbook workbook = new Workbook(outputStream, "workbook","1.0");
		Worksheet firstSheet = workbook.newWorksheet("Sheet1");
		insertRow(firstSheet, xls.getHeaders(), 0);
		int rowNum = 1;
		if (xls.getData().size() > 0) {
			for (List<String> rowData : xls.getData()) {
				insertRow(firstSheet, rowData, rowNum);
				rowNum++;
			}
		}
		workbook.finish();
		return outputStream;
	}
	public Workbook listsToWorkbook(List<List<String>> companyData) throws IOException {
		Integer listSize = companyData.size();
		List<String> headers = companyData.get(listSize - 1);
		companyData.remove(listSize - 1);
		StableDocumentXls xls = StableDocumentXls.builder().build();
		xls.setHeaders(headers);
		xls.setData(companyData);
		Workbook workbook = writeToWorkBook(xls);
		return workbook;
	}




	/**
	 * Convert the size of the document. The
	 * subDocumentXls1 begins at the specified {@code beginIndex} and
	 * extends to the character at index {@code endIndex - 1}.
	 * Thus the length of the subDocumentXls1 is {@code endIndex-beginIndex}.
	 * <p>
	 * Examples:
	 * <blockquote><pre>
	 * "hamburger".subDocumentXls1(4, 8) returns "urge"
	 * "smiles".subDocumentXls1(1, 5) returns "mile"
	 * </pre></blockquote>
	 *
	 * @param      beginIndex   the beginning index, inclusive.
	 * @param      endIndex     the ending index, exclusive.
	 * @return     the specified subDocumentXls1.
	 * @exception  IndexOutOfBoundsException  if the
	 *             {@code beginIndex} is negative, or
	 *             {@code endIndex} is larger than the length of
	 *             this {@code DocumentXls1} object, or
	 *             {@code beginIndex} is larger than
	 *             {@code endIndex}.
	 */
	public void subDocumentXls1(StableDocumentXls xls, int beginIndex, int endIndex) {
		if (xls.getData().isEmpty() ||beginIndex < 0 || endIndex > xls.getData().size())
			return;

		List<List<String>> newData = new ArrayList<>();

		for (int rowIndex = beginIndex; rowIndex < endIndex; rowIndex++) {
			List<String> rowData = xls.getData().get(rowIndex);
			newData.add(rowData);
		}

		xls.setData(newData);
	}


	public int createColumn(StableDocumentXls stableDocumentXls, String columnName) {
		int index = stableDocumentXls.getHeaders().indexOf(columnName);
		if (index == -1) {
			index = stableDocumentXls.getHeaders().size();
			stableDocumentXls.getHeaders().add(index, columnName);
			stableDocumentXls.getIndexColumnMap().put(index,columnName);
			stableDocumentXls.getColumnIndexMap().put(columnName, index);

		}

		if (stableDocumentXls.getData() == null)
			stableDocumentXls.setData(new ArrayList<>());
		for (int rowIndex = 0; rowIndex < stableDocumentXls.getData().size(); rowIndex++) {
			List<String> row = stableDocumentXls.getData().get(rowIndex);
//			Map<String, String> rowByHeaderName = stableDocumentXls.getDataByHeaderName().get(rowIndex);
			for (int left = row.size(); left <= index; left++)
				row.add(left, "");

		}
		return index;
	}

}


