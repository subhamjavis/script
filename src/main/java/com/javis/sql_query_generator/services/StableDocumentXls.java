package com.javis.sql_query_generator.services;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@AllArgsConstructor
@NoArgsConstructor
@Builder(toBuilder=true)
@Data
public class StableDocumentXls {

	private String fileName;

	private String keyValuePair;

	private List<String> headers = new ArrayList<>();

	private Map<String, String> headerDataType = new HashMap<>();

	private List<List<String>> data = new ArrayList<>();

	private Map<String,Integer> columnIndexMap = new HashMap<>();
	private Map<Integer, String> indexColumnMap = new HashMap<>();

//	private List<Map<String, String>> dataByHeaderName = new ArrayList<>();

	public Map<String,Integer> getColumnIndexMap(){ return this.columnIndexMap; }
	public Map<Integer, String> getIndexColumnMap(){ return  this.indexColumnMap; }

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getKeyValuePair() {
		return keyValuePair;
	}

	public void setKeyValuePair(String keyValuePair) {
		this.keyValuePair = keyValuePair;
	}

	public List<String> getHeaders() {
		return headers;
	}

	public void setHeaders(List<String> headers) {
		this.headers = headers;
	}

	public void setColumnIndexMap(Map<String, Integer> columnIndexMap){ this.columnIndexMap = columnIndexMap ;}
	public void setIndexColumnMap(Map<Integer, String> indexColumnMap){ this.indexColumnMap = indexColumnMap; }

	public Map<String, String> getHeaderDataType() {
		return headerDataType;
	}

	public void setHeaderDataType(Map<String, String> headerDataType) {
		this.headerDataType = headerDataType;
	}

	public List<List<String>> getData() {
		return data;
	}

	public void setData(List<List<String>> data) {
		this.data = data;
	}

//	public List<Map<String, String>> getDataByHeaderName() {
//		return dataByHeaderName;
//	}
//
//	public void setDataByHeaderName(List<Map<String, String>> dataByHeaderName) {
//		this.dataByHeaderName = dataByHeaderName;
//	}

	@Override
	public String toString() {
		return "DocumentXls [fileName=" + fileName + ", keyValuePair=" + keyValuePair + ", headers=" + headers
				+ ", headerDataType=" + headerDataType + ", length=" + data.size() + "]";
	}

}
